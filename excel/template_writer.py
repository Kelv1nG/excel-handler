from __future__ import annotations

import re
from copy import copy
from os import PathLike

from openpyxl.utils import coordinate_to_tuple

from typing import Any

import polars as pl

from excel.protocols import TypedValue
from excel.template_reader import ExcelTemplateReader, MarkedCell
from excel._utils import load_excel_workbook

_END_TABLE_RE = re.compile(r"\{\{\s*end_table\s*(?:\|\s*(?P<meta>[^}]*))?\}\}")
_INSERT_DATA_RE = re.compile(r"\{\{\s*insert_data\s*\}\}")

def _is_loop(cell: MarkedCell) -> bool:
    """Return ``True`` if *cell* carries a ``loop()`` metadata tag."""
    return cell.parse_metadata().get("type") == "loop"


def _is_table(cell: MarkedCell) -> bool:
    """Return ``True`` if *cell* carries a ``table()`` metadata tag."""
    return cell.parse_metadata().get("type") == "table"


def _parse_table_meta(cell: MarkedCell) -> tuple[str, str | None]:
    """Return (join_mode, on_col_override_or_None). Default join mode is 'left'."""
    meta = cell.parse_metadata()
    join_mode = str(meta.get("join", "left"))
    on_col = str(meta["on"]) if "on" in meta else None
    return join_mode, on_col


def _get_style_source(ws, row: int, col: int):
    """Return the cell that owns the style for *(row, col)*.

    Inside a merged range only the top-left cell carries style data;
    every other cell is a MergedCell proxy with ``has_style == False``.
    This helper resolves to the top-left cell when needed.
    """
    cell = ws.cell(row, col)
    if cell.has_style:
        return cell
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(m.min_row, m.min_col)
    return cell


def _get_merged_cell_value(ws, row: int, col: int) -> Any:
    """Return the effective value of a cell, resolving merged ranges.

    In normal (non-read-only) mode openpyxl returns ``None`` for every cell
    inside a merged range except the top-left corner.  This helper checks
    whether *(row, col)* belongs to a merged range and, if so, returns the
    value stored in the top-left cell of that range.
    """
    val = ws.cell(row, col).value
    if val is not None:
        return val
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            return ws.cell(m.min_row, m.min_col).value
    return None


def _copy_row_styles(ws, source_row: int, count: int) -> None:
    """Insert *count* rows below *source_row*, copying values and styles from it."""
    # Snapshot merged ranges before insert — openpyxl's insert_rows() auto-extends
    # any merged range that spans the insertion point, which would incorrectly
    # merge the newly inserted data rows.
    saved_merges = [
        (m.min_row, m.min_col, m.max_row, m.max_col)
        for m in ws.merged_cells.ranges
    ]

    # Snapshot source-row styles BEFORE insert_rows() — the insert can create
    # phantom MergedCell proxies on source_row that the later ghost purge
    # removes, leaving fresh Cell objects with no styles.  By saving a copy
    # now we can restore source_row and copy to inserted rows reliably.
    max_col = ws.max_column
    saved_styles: dict[int, dict[str, Any]] = {}
    for col in range(1, max_col + 1):
        style_cell = _get_style_source(ws, source_row, col)
        if style_cell.has_style:
            saved_styles[col] = {
                "font": copy(style_cell.font),
                "border": copy(style_cell.border),
                "fill": copy(style_cell.fill),
                "number_format": style_cell.number_format,
                "protection": copy(style_cell.protection),
                "alignment": copy(style_cell.alignment),
            }

    ws.insert_rows(source_row + 1, count)

    # Undo openpyxl's automatic adjustments and re-apply with correct logic:
    #   - Ranges entirely at or above source_row → unchanged
    #   - Ranges entirely below source_row → shift down by count
    #   - Ranges spanning source_row → split into top / bottom halves,
    #     leaving the newly inserted rows unmerged
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    for min_r, min_c, max_r, max_c in saved_merges:
        if max_r <= source_row:
            ws.merge_cells(
                start_row=min_r, start_column=min_c,
                end_row=max_r, end_column=max_c,
            )
        elif min_r > source_row:
            ws.merge_cells(
                start_row=min_r + count, start_column=min_c,
                end_row=max_r + count, end_column=max_c,
            )
        else:
            # Top portion (min_r … source_row) — keep only if multi-cell
            if source_row > min_r or max_c > min_c:
                ws.merge_cells(
                    start_row=min_r, start_column=min_c,
                    end_row=source_row, end_column=max_c,
                )
            # Bottom portion (rows that were below source_row, now shifted)
            if max_r > source_row:
                bottom_start = source_row + count + 1
                bottom_end = max_r + count
                if bottom_end > bottom_start or max_c > min_c:
                    ws.merge_cells(
                        start_row=bottom_start, start_column=min_c,
                        end_row=bottom_end, end_column=max_c,
                    )

    # Safety: ensure no merge overlaps the inserted rows.  openpyxl can leave
    # phantom MergedCell objects when insert_rows touches an existing merge
    # boundary, causing cells to appear merged in the output.
    inserted_start = source_row + 1
    inserted_end = source_row + count
    for m in list(ws.merged_cells.ranges):
        if m.min_row <= inserted_end and m.max_row >= inserted_start:
            ws.unmerge_cells(str(m))

    # After unmerging, openpyxl leaves MergedCell ghost objects in ws._cells.
    # These silently ignore style writes (.font, .alignment, etc.).  Purge
    # them so the next ws.cell() call creates fresh, writable Cell objects.
    # Include source_row: insert_rows() can create phantom merges that touch
    # it, converting real Cells to MergedCell proxies and losing their values.
    from openpyxl.cell.cell import MergedCell as _MC
    for r in range(source_row, inserted_end + 1):
        for c in list(ws._cells):
            if isinstance(ws._cells.get(c), _MC) and c[0] == r:
                del ws._cells[c]

    # Restore source_row styles (may have been wiped by ghost purge) and
    # copy styles + values to each inserted row.
    for col in range(1, max_col + 1):
        src_cell = ws.cell(source_row, col)
        style = saved_styles.get(col)
        if style:
            src_cell.font = copy(style["font"])
            src_cell.border = copy(style["border"])
            src_cell.fill = copy(style["fill"])
            src_cell.number_format = style["number_format"]
            src_cell.protection = copy(style["protection"])
            src_cell.alignment = copy(style["alignment"])

    for offset in range(1, count + 1):
        for col in range(1, max_col + 1):
            val_src = ws.cell(source_row, col)
            dst = ws.cell(source_row + offset, col)
            dst.value = val_src.value
            style = saved_styles.get(col)
            if style:
                dst.font = copy(style["font"])
                dst.border = copy(style["border"])
                dst.fill = copy(style["fill"])
                dst.number_format = style["number_format"]
                dst.protection = copy(style["protection"])
                dst.alignment = copy(style["alignment"])


def _read_headers(ws, header_row: int, start_col: int) -> list[tuple[str, int]]:
    """Read non-empty header names rightward from start_col.

    Returns list of (column_name, col_index) pairs.
    """
    headers: list[tuple[str, int]] = []
    col = start_col
    while True:
        val = _get_merged_cell_value(ws, header_row, col)
        if val is None or str(val).strip() == "":
            break
        headers.append((str(val), col))
        col += 1
    return headers


def _find_last_data_row(
    ws,
    start_row: int,
    join_col: int,
    data_cols: list[int] | None = None,
) -> int:
    """Return the last row in join_col with a non-empty value, starting from start_row.

    When *data_cols* is provided the scan also requires at least one of those
    columns to be non-empty — this prevents stray text below the table (e.g.
    footnotes that happen to sit in the join column) from being treated as data
    rows.  The very first row (``start_row``) is always accepted when its join
    column is non-empty, because it is the tag row and its data columns may
    contain the ``{{ tag }}`` placeholder which has already been cleared.
    """
    last_row = start_row
    row = start_row
    while True:
        val = _get_merged_cell_value(ws, row, join_col)
        if val is None or str(val).strip() == "":
            break
        # For rows after start_row, verify at least one data column is filled.
        if data_cols and row > start_row:
            has_data = any(
                _get_merged_cell_value(ws, row, c) not in (None, "")
                for c in data_cols
            )
            if not has_data:
                break
        last_row = row
        row += 1
    return last_row


def _find_end_table_row(ws, start_row: int) -> tuple[int, int, str] | None:
    """Scan rows from *start_row* downward for a ``{{ end_table }}`` marker.

    Returns ``(row, col, insert_mode)`` of the marker cell, or ``None`` if
    not found.  *insert_mode* is ``"below"`` (default) or ``"above"``.
    Stops at the first fully-blank row (all columns empty).
    """
    max_col = ws.max_column or 1
    row = start_row
    while True:
        all_empty = True
        for col in range(1, max_col + 1):
            val = ws.cell(row, col).value
            if val is not None and str(val).strip():
                all_empty = False
                if isinstance(val, str):
                    m = _END_TABLE_RE.search(val)
                    if m:
                        insert_mode = "below"
                        meta = (m.group("meta") or "").strip()
                        if meta:
                            for part in meta.split(","):
                                k, _, v = part.partition("=")
                                if k.strip().lower() == "insert":
                                    insert_mode = v.strip().lower()
                        return (row, col, insert_mode)
        if all_empty:
            break
        row += 1
    return None


def _find_insert_data_row(ws, start_row: int) -> tuple[int, int] | None:
    """Scan rows from *start_row* downward for a ``{{ insert_data }}`` marker.

    Returns ``(row, col)`` of the marker cell, or ``None`` if not found.
    Stops at the first fully-blank row.
    """
    max_col = ws.max_column or 1
    row = start_row
    while True:
        all_empty = True
        for col in range(1, max_col + 1):
            val = ws.cell(row, col).value
            if val is not None and str(val).strip():
                all_empty = False
                if isinstance(val, str) and _INSERT_DATA_RE.search(val):
                    return (row, col)
        if all_empty:
            break
        row += 1
    return None


def _fill_table(ws, mc: MarkedCell, df: pl.DataFrame) -> None:
    """Fill ws with df data using join semantics described in mc.metadata.

    Join modes:
    - left:  fill matched rows; leave unmatched template rows blank. No inserts.
    - inner: fill matched rows; clear data cols on unmatched template rows. No inserts.
    - outer: fill matched rows; append unmatched DF rows at bottom (pushes content down).
    - right: overwrite template rows top-down in DF order; insert if DF is longer;
             clear remaining template rows if DF is shorter.
    """
    join_mode, on_col = _parse_table_meta(mc)
    tag_row, tag_col = coordinate_to_tuple(mc.cell_addr)
    header_row = tag_row - 1
    join_tmpl_col = tag_col - 1

    # Name of the DF join column: explicit on= override, else template header name
    tmpl_join_header = _get_merged_cell_value(ws, header_row, join_tmpl_col)
    join_df_col: str = on_col if on_col is not None else str(tmpl_join_header)

    headers = _read_headers(ws, header_row, tag_col)
    data_col_indices = [col_idx for _, col_idx in headers]

    # Handle {{ insert_data }} marker — delete its row early so the boundary
    # scans below see a contiguous table.  The deleted row's position becomes
    # the insertion point for extra rows (outer join).
    insert_data_marker = _find_insert_data_row(ws, tag_row)
    insert_data_before: int | None = None
    if insert_data_marker is not None:
        id_row, _id_col = insert_data_marker
        ws.delete_rows(id_row)
        # After deletion, the row that was below now sits at id_row.
        # Extra rows should be inserted above that row.
        insert_data_before = id_row

    # Determine table boundary.  An explicit {{ end_table }} marker takes
    # precedence over the heuristic scan in _find_last_data_row.
    #
    # Three modes:
    #   Option A: {{ end_table }} on its own row (no join value) — table ends
    #             at the row above; marker row is deleted after writes.
    #   Option B: {{ end_table }} on a data row — marker cell cleared; that
    #             row is the last table row.  Extras go after it.
    #   Option C: {{ end_table | insert=above }} on a data row — marker cell
    #             cleared; that row is part of the table AND gets matched,
    #             but extras are inserted BEFORE it (insert_before_row).
    end_marker = _find_end_table_row(ws, tag_row)
    delete_end_row: int | None = None
    insert_before_row: int | None = None  # Option C: extras go before this row

    if end_marker is not None:
        end_row, end_col, insert_mode = end_marker
        end_join_val = _get_merged_cell_value(ws, end_row, join_tmpl_col)
        if end_join_val is not None and str(end_join_val).strip():
            # Data row with {{ end_table }}.  Clear the marker cell.
            ws.cell(end_row, end_col).value = None
            if insert_mode == "above":
                # Option C: match all rows through end_row, but insert
                # extras above end_row (between the data rows and this row).
                last_tmpl_row = end_row
                insert_before_row = end_row
            else:
                # Option B: extras go after last_tmpl_row (default).
                last_tmpl_row = end_row
        else:
            # Option A: {{ end_table }} on its own row.
            last_tmpl_row = end_row - 1
            delete_end_row = end_row
    else:
        last_tmpl_row = _find_last_data_row(ws, tag_row, join_tmpl_col, data_col_indices)

    # {{ insert_data }} takes precedence over {{ end_table }} for the
    # insertion point.  The table boundary (last_tmpl_row) is unaffected —
    # matching continues through all rows.
    if insert_data_before is not None:
        insert_before_row = insert_data_before

    n_tmpl = last_tmpl_row - tag_row + 1

    # Track how many rows were inserted so we can adjust the delete_end_row
    # position at the end.
    rows_inserted = 0

    # Clear the tag placeholder before any data writes so the data loop
    # can overwrite (tag_row, tag_col) with the correct value.
    ws.cell(tag_row, tag_col).value = None

    if join_mode == "right":
        df_list = df.to_dicts()
        n_df = len(df_list)
        if n_df > n_tmpl:
            rows_inserted = n_df - n_tmpl
            _copy_row_styles(ws, last_tmpl_row, rows_inserted)
        for i, row in enumerate(df_list):
            ws_row = tag_row + i
            ws.cell(ws_row, join_tmpl_col).value = row.get(join_df_col)
            for col_name, col_idx in headers:
                ws.cell(ws_row, col_idx).value = row.get(col_name)
        # Clear extra template rows if DF has fewer rows
        for r in range(tag_row + n_df, last_tmpl_row + 1):
            ws.cell(r, join_tmpl_col).value = None
            for _, col_idx in headers:
                ws.cell(r, col_idx).value = None
    else:
        # Build DF lookup: join_value → row dict
        df_lookup: dict[Any, dict[str, Any]] = {}
        for row in df.iter_rows(named=True):
            df_lookup[row[join_df_col]] = row

        # Snapshot per-row join values BEFORE any row insertion.  _copy_row_styles
        # may purge MergedCell ghosts on source_row, wiping the join column
        # value and causing the matched-row lookup to miss that row entirely.
        saved_join: dict[int, Any] = {
            r: _get_merged_cell_value(ws, r, join_tmpl_col)
            for r in range(tag_row, last_tmpl_row + 1)
        }

        # outer: insert extra rows BEFORE writing any data.  insert_rows()
        # can create phantom MergedCell proxies on source_row, silently
        # destroying values already written there.  By inserting first and
        # writing matched rows afterwards, all cells are real Cell objects.
        extra: list[dict[str, Any]] = []
        if join_mode == "outer":
            tmpl_vals = set(saved_join.values())
            extra = [
                row
                for row in df.iter_rows(named=True)
                if row[join_df_col] not in tmpl_vals
            ]
            if extra:
                rows_inserted = len(extra)
                if insert_before_row is not None:
                    # Option C: insert extras BEFORE the insert_before_row.
                    # Use the row above as the style source.
                    style_src = insert_before_row - 1
                    if style_src < tag_row:
                        style_src = tag_row
                    _copy_row_styles(ws, style_src, rows_inserted)
                    # Rows were inserted after style_src, so
                    # insert_before_row (and everything below) shifted down.
                    insert_before_row += rows_inserted
                    last_tmpl_row += rows_inserted
                else:
                    _copy_row_styles(ws, last_tmpl_row, rows_inserted)

        # After row insertion, saved_join keys may no longer match worksheet
        # row numbers.  Rebuild the mapping: original template rows that were
        # at or above the insertion point keep their row number; rows at or
        # below insert_before_row (Option C) have shifted down by
        # rows_inserted.
        if rows_inserted and insert_before_row is not None:
            shifted_join: dict[int, Any] = {}
            # insert_before_row already accounts for the shift (updated above).
            # Original rows were at: insert_before_row - rows_inserted .. last_tmpl_row - rows_inserted
            orig_insert_row = insert_before_row - rows_inserted
            for orig_r, val in saved_join.items():
                if orig_r >= orig_insert_row:
                    shifted_join[orig_r + rows_inserted] = val
                else:
                    shifted_join[orig_r] = val
            saved_join = shifted_join

        # Fill / clear matched template rows (safe now — any row insertion
        # and ghost purge has already happened above).  Use saved_join
        # instead of re-reading from the worksheet, because the ghost purge
        # inside _copy_row_styles may have wiped the join column cell.
        for r in sorted(saved_join):
            tmpl_val = saved_join[r]
            if tmpl_val in df_lookup:
                row = df_lookup[tmpl_val]
                for col_name, col_idx in headers:
                    ws.cell(r, col_idx).value = row.get(col_name)
            elif join_mode == "inner":
                for _, col_idx in headers:
                    ws.cell(r, col_idx).value = None
            # left: leave unmatched rows as-is

        # Write extra rows (outer join only).
        if extra:
            if insert_before_row is not None:
                # Option C: extras occupy the rows just above insert_before_row.
                extra_start = insert_before_row - rows_inserted
            else:
                extra_start = last_tmpl_row + 1
            for i, row in enumerate(extra):
                ws_row = extra_start + i
                ws.cell(ws_row, join_tmpl_col).value = row[join_df_col]
                for col_name, col_idx in headers:
                    ws.cell(ws_row, col_idx).value = row.get(col_name)

    # Delete the {{ end_table }} marker row (Option A).  Row indices may have
    # shifted if extra rows were inserted above, so adjust accordingly.
    if delete_end_row is not None:
        ws.delete_rows(delete_end_row + rows_inserted)


class ExcelTemplateWriter:
    """Fill an Excel template with data and write the output file.

    Template syntax:

    * ``{{ variable }}`` — replaced with the scalar value of *variable*.
    * ``{{ variable | loop }}`` — marks the cell as part of a loop row.
      All loop-tagged cells in the same row must reference variables whose
      values are lists of the same length.  The template row is expanded
      into N rows (one per list element).

    Usage::

        writer = ExcelTemplateWriter("template.xlsx")
        writer.write(
            {
                "title": TypedValue("Q1 Report", "single"),
                "month": TypedValue(["Jan", "Feb", "Mar"], "list"),
                "amount": TypedValue([100, 200, 300], "list"),
            },
            "output.xlsx",
        )
    """

    def __init__(self, template: str | PathLike[str] | bytes):
        self._template = template

    def write(self, vars: dict[str, TypedValue], file: str | PathLike[str]) -> None:
        """Fill the template and save to *file*.

        Args:
            vars: Variable name → TypedValue.
            file: Output path for the filled workbook.

        Raises:
            KeyError: A template tag references a variable not present in *vars*.
            ValueError: Loop cells in the same row have lists of different lengths.
        """
        wb = load_excel_workbook(self._template)
        structure = ExcelTemplateReader().read(self._template)

        # Split tagged cells into loop participants, tables, and plain scalars.
        loop_rows: dict[tuple[str, int], list[MarkedCell]] = {}
        tables: list[MarkedCell] = []
        scalars: list[MarkedCell] = []

        for sheet_name, cells in structure.items():
            for cell in cells:
                if cell.name in ("end_table", "insert_data"):
                    continue  # structural marker, not a variable
                if _is_table(cell):
                    tables.append(cell)
                elif _is_loop(cell):
                    row_n = coordinate_to_tuple(cell.cell_addr)[0]
                    loop_rows.setdefault((sheet_name, row_n), []).append(cell)
                else:
                    scalars.append(cell)

        # ── Loop rows ──────────────────────────────────────────────────────────
        # Group by sheet then process each sheet's rows bottom-up so row inserts
        # don't invalidate later row indices within the same sheet.
        by_sheet: dict[str, list[tuple[int, list[MarkedCell]]]] = {}
        for (sheet_name, row_n), cells in loop_rows.items():
            by_sheet.setdefault(sheet_name, []).append((row_n, cells))

        for sheet_name, rows in by_sheet.items():
            ws = wb[sheet_name]
            for row_n, cells in sorted(rows, key=lambda x: x[0], reverse=True):
                n = len(vars[cells[0].name].value)

                for mc in cells[1:]:
                    actual = len(vars[mc.name].value)
                    if actual != n:
                        raise ValueError(
                            f"Loop variables in row {row_n} of '{sheet_name}' have "
                            f"different lengths: '{cells[0].name}'={n}, "
                            f"'{mc.name}'={actual}"
                        )

                if n == 0:
                    ws.delete_rows(row_n)
                    continue

                if n > 1:
                    _copy_row_styles(ws, row_n, n - 1)

                for i in range(n):
                    for mc in cells:
                        _, col_n = coordinate_to_tuple(mc.cell_addr)
                        ws.cell(row_n + i, col_n).value = vars[mc.name].value[i]

        # ── Table cells ────────────────────────────────────────────────────────
        # Process bottom-up per sheet so outer row insertions don't corrupt
        # later table positions in the same sheet.
        tables_by_sheet: dict[str, list[MarkedCell]] = {}
        for mc in tables:
            tables_by_sheet.setdefault(mc.sheet, []).append(mc)

        for sheet_name, mcs in tables_by_sheet.items():
            ws = wb[sheet_name]
            for mc in sorted(
                mcs, key=lambda c: coordinate_to_tuple(c.cell_addr)[0], reverse=True
            ):
                _fill_table(ws, mc, vars[mc.name].value)

        # ── Scalar cells ───────────────────────────────────────────────────────
        for mc in scalars:
            ws = wb[mc.sheet]
            row_n, col_n = coordinate_to_tuple(mc.cell_addr)
            ws.cell(row_n, col_n).value = vars[mc.name].value

        wb.save(file)
