"""Type definitions and metadata models for template filling.

Exports:
    JoinMode, InsertMode, StyleSource — Literal type aliases
    OrderBy, FillSpec, TableMeta, ImageMeta — dataclasses for parsing template tag metadata
    _EndTableMarker — structural marker for {{ end_table }} cells
    _apply_fill, _is_loop, _is_table, _is_image — helper functions
    _FILL_MISSING — sentinel for missing fill spec keys
    _TemplateRegex — regex patterns for metadata extraction
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Literal, cast

from excel.template_reader import MarkedCell

JoinMode = Literal["left", "inner", "outer", "right"]
InsertMode = Literal["above", "below"]
StyleSource = Literal["first", "last"]

_VALID_JOIN_MODES: frozenset[str] = frozenset({"left", "inner", "outer", "right"})
_VALID_STYLE_SOURCES: frozenset[str] = frozenset({"first", "last"})

_FILL_MISSING = object()  # sentinel — "key was not present in fill spec"


def _coerce_fill(v: str) -> Any:
    """Coerce a fill value string to its most specific Python type."""
    if v.lower() == "true":
        return True
    if v.lower() == "false":
        return False
    try:
        return int(v)
    except ValueError:
        pass
    try:
        return float(v)
    except ValueError:
        pass
    return v


@dataclass(frozen=True)
class OrderBy:
    """Represents an ``order_by`` sort directive on an outer-join table tag.

    Attributes:
        col:       DataFrame column to sort by.  ``None`` means "use the join column".
        ascending: ``True`` for ascending, ``False`` for descending.
    """

    col: str | None
    ascending: bool

    @classmethod
    def from_meta(cls, raw: str) -> OrderBy:
        """Parse the value of an ``order_by=`` metadata key.

        Accepted forms:

        * ``asc`` / ``desc``                    → sort by the join column
        * ``ColName``                           → sort by *ColName*, ascending
        * ``ColName:asc`` / ``ColName:desc``    → sort by *ColName* with explicit direction
        """
        v = raw.strip()
        if v.lower() == "asc":
            return cls(col=None, ascending=True)
        if v.lower() == "desc":
            return cls(col=None, ascending=False)
        if ":" in v:
            col, _, direction = v.partition(":")
            return cls(col=col.strip() or None, ascending=direction.strip().lower() != "desc")
        return cls(col=v or None, ascending=True)


class _TemplateRegex:
    """Regex patterns for parsing template metadata."""
    END_TABLE = re.compile(r"\{\{\s*end_table\s*(?:\|\s*(?P<meta>[^}]*))?\}\}")
    INSERT_DATA = re.compile(r"\{\{\s*insert_data\s*\}\}")
    # Extracts the raw fill spec from the metadata string without relying on the
    # comma-split parser.  Matches "fill=..." stopping before the next key= or
    # closing paren, e.g. "fill=0" or "fill=colA:0;colB:N/A".
    FILL_SPEC = re.compile(r"\bfill=([^,)]+)")


@dataclass(frozen=True)
class FillSpec:
    """Represents a ``fill=`` directive that substitutes ``None`` values in data columns.

    Attributes:
        values: Mapping of column name → fill value.  The key ``None`` is a global
                fallback applied when no per-column entry exists.

    Examples::

        FillSpec({None: 0})                    # global: replace every None with 0
        FillSpec({"ColA": 0, "ColB": "N/A"})   # per-column fills
    """

    values: dict[str | None, Any]

    def apply(self, value: Any, col_name: str) -> Any:
        """Return *value*, or a fill substitute when *value* is ``None``.

        Lookup order: per-column entry → global (``None`` key) → unchanged ``None``.
        """
        if value is not None:
            return value
        v = self.values.get(col_name, _FILL_MISSING)
        if v is not _FILL_MISSING:
            return v
        return self.values.get(None)  # global fallback; None if absent

    @classmethod
    def from_meta(cls, metadata: str) -> FillSpec | None:
        """Parse a ``fill=`` metadata parameter, returning ``None`` when absent.

        Two forms:

        * ``fill=0``                → global: ``FillSpec({None: 0})``
        * ``fill=colA:0;colB:N/A`` → per-column (colon-separated col:value pairs,
          semicolons between pairs — commas are reserved by the outer metadata parser)
        """
        m = _TemplateRegex.FILL_SPEC.search(metadata)
        if m is None:
            return None
        raw = m.group(1).strip()
        if ":" not in raw:
            return cls(values={None: _coerce_fill(raw)})
        result: dict[str | None, Any] = {}
        for part in raw.split(";"):
            part = part.strip()
            if not part:
                continue
            col, _, val = part.partition(":")
            result[col.strip()] = _coerce_fill(val.strip())
        return cls(values=result)


@dataclass(frozen=True)
class TableMeta:
    """Parsed metadata for a ``{{ variable | table(...) }}`` tag.

    Attributes:
        join:        Join strategy — ``"left"``, ``"inner"``, ``"outer"``, or ``"right"``.
        on:          Explicit join-column override.  ``None`` means "infer from header".
        positional:  When ``True``, fill by position with no join column.
        placeholder: When ``True`` (outer only), the tag row is used as a style source
                     then deleted if its join value is absent from the DataFrame.
        style:       Which template row to copy styles from when inserting rows —
                     ``"first"`` (tag row) or ``"last"`` (last template row, default).
        order_by:    Sort directive for the upper zone in outer joins.  ``None`` = unsorted.
        fill:        Fill spec substituting ``None`` values in data columns.  ``None`` = no fill.
    """

    join: JoinMode
    on: str | None
    positional: bool
    placeholder: bool
    style: StyleSource
    order_by: OrderBy | None
    fill: FillSpec | None

    @classmethod
    def from_cell(cls, mc: MarkedCell) -> TableMeta:
        """Build a ``TableMeta`` from a tagged ``MarkedCell``.

        Raises:
            ValueError: ``style=`` value is not ``"first"`` or ``"last"``.
            ValueError: ``join=`` value is not a recognised join mode.
        """
        meta = mc.parse_metadata()

        raw_join = str(meta.get("join", "left"))
        if raw_join not in _VALID_JOIN_MODES:
            raise ValueError(
                f"Invalid join={raw_join!r} in tag {mc.raw!r}. "
                f"Expected one of: {', '.join(sorted(_VALID_JOIN_MODES))}."
            )

        raw_style = str(meta.get("style", "last"))
        if raw_style not in _VALID_STYLE_SOURCES:
            raise ValueError(
                f"Invalid style={raw_style!r} in tag {mc.raw!r}. "
                "Expected 'first' or 'last'."
            )

        order_by: OrderBy | None = None
        raw_order = meta.get("order_by")
        if raw_order is not None:
            order_by = OrderBy.from_meta(str(raw_order))

        return cls(
            join=cast(JoinMode, raw_join),
            on=str(meta["on"]) if "on" in meta else None,
            positional=bool(meta.get("positional", False)),
            placeholder=bool(meta.get("placeholder", False)),
            style=cast(StyleSource, raw_style),
            order_by=order_by,
            fill=FillSpec.from_meta(mc.metadata),
        )


def _apply_fill(value: Any, col_name: str, fill: FillSpec | None) -> Any:
    """Return *value*, or a fill substitute when *value* is ``None`` and *fill* is set."""
    return fill.apply(value, col_name) if fill is not None else value


def _is_loop(cell: MarkedCell) -> bool:
    """Return ``True`` if *cell* carries a ``loop()`` metadata tag."""
    return cell.parse_metadata().get("type") == "loop"


def _is_table(cell: MarkedCell) -> bool:
    """Return ``True`` if *cell* carries a ``table()`` metadata tag."""
    return cell.parse_metadata().get("type") == "table"


def _is_image(cell: MarkedCell) -> bool:
    """Return ``True`` if *cell* carries an ``image()`` metadata tag."""
    return cell.parse_metadata().get("type") == "image"


@dataclass(frozen=True)
class ImageMeta:
    """Parsed metadata for a ``{{ variable | image(...) }}`` tag.

    Attributes:
        width:  Override width in pixels.  ``None`` keeps the image's natural width.
        height: Override height in pixels.  ``None`` keeps the image's natural height.
    """

    width: int | None
    height: int | None

    @classmethod
    def from_cell(cls, mc: MarkedCell) -> ImageMeta:
        """Build an ``ImageMeta`` from a tagged ``MarkedCell``."""
        meta = mc.parse_metadata()
        raw_w = meta.get("width")
        raw_h = meta.get("height")
        return cls(
            width=int(raw_w) if raw_w is not None else None,
            height=int(raw_h) if raw_h is not None else None,
        )


@dataclass(frozen=True)
class _EndTableMarker:
    """Location and insert-mode of a ``{{ end_table }}`` structural marker cell.

    Attributes:
        row:         1-based worksheet row of the marker cell.
        col:         1-based worksheet column of the marker cell.
        insert_mode: Where extra outer-join rows are placed relative to the marker —
                     ``"below"`` (default) or ``"above"``.
    """

    row: int
    col: int
    insert_mode: InsertMode
