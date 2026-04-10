# Table Template Examples

This document shows every table tag feature with a concrete template layout, input DataFrame, and the resulting output. All examples use the convention:

- `[ ]` = cell in template
- Greyed values are pre-existing template content
- `↓ N rows inserted` annotations mark rows the writer adds at runtime

---

## Contents

1. [Left join (default)](#1-left-join-default)
2. [Inner join](#2-inner-join)
3. [Outer join — fewer template rows than data](#3-outer-join--fewer-template-rows-than-data)
4. [Outer join — more template rows than data](#4-outer-join--more-template-rows-than-data)
5. [Right join](#5-right-join)
6. [Boundary markers: `end_table`](#6-boundary-markers-end_table)
7. [Boundary markers: `end_table | insert=above` (pin a row)](#7-boundary-markers-end_table--insertabove-pin-a-row)
8. [Boundary markers: `insert_data`](#8-boundary-markers-insert_data)
9. [Pinned Total row with `placeholder=true`](#9-pinned-total-row-with-placeholdertrue)
10. [Sorted outer join (`order_by`)](#10-sorted-outer-join-order_by)
11. [Fill missing values (`fill=`)](#11-fill-missing-values-fill)
12. [Style source for inserted rows (`style=first|last`)](#12-style-source-for-inserted-rows-stylefirstlast)
13. [Positional fill (`positional=true`)](#13-positional-fill-positionaltrue)
14. [Custom join column (`on=`)](#14-custom-join-column-on)
15. [Combining options](#15-combining-options)

---

## 1. Left join (default)

The writer finds each template row's join key in the DataFrame and fills the data columns. Template rows whose key is **not** in the DataFrame are left as-is. No rows are inserted or deleted.

**Tag**
```
{{ sales | table() }}
```
*(default: `join=left`)*

**Template**

| Region | Revenue | Margin |
|--------|---------|--------|
| North  | `{{ sales \| table() }}` | |
| South  | | |
| East   | | |
| West   | | |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East"],
    "Revenue": [1000, 800, 600],
    "Margin":  [0.20, 0.15, 0.12],
})
```

**Output**

| Region | Revenue | Margin |
|--------|---------|--------|
| North  | 1000 | 0.20 |
| South  | 800  | 0.15 |
| East   | 600  | 0.12 |
| West   | *(unchanged)* | *(unchanged)* |

> West has no match → its data columns are left exactly as they are in the template (usually blank, or a pre-filled default).

---

## 2. Inner join

Like left join, but template rows whose key is **not** in the DataFrame have their data columns **cleared**.

**Tag**
```
{{ sales | table(join=inner) }}
```

**Template** *(same as Example 1)*

**DataFrame** *(same as Example 1 — no West)*

**Output**

| Region | Revenue | Margin |
|--------|---------|--------|
| North  | 1000 | 0.20 |
| South  | 800  | 0.15 |
| East   | 600  | 0.12 |
| West   | *(cleared)* | *(cleared)* |

> West survives as a row (no deletion), but its Revenue and Margin cells are set to `None`.

---

## 3. Outer join — fewer template rows than data

Extra DataFrame rows (those whose key is absent from the template) are **inserted** after the last template data row. Everything below shifts down.

**Tag**
```
{{ sales | table(join=outer) }}
```

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | (headers) | |
| 2   | North  | `{{ sales \| table(join=outer) }}` |
| 3   | South  | |
| 4   | *(footer, merged A4:B4)* | |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East", "West"],
    "Revenue": [1000, 800, 600, 400],
})
```

**Output**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | (headers) | |
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  | ← inserted
| 5   | West   | 400  | ← inserted
| 6   | *(footer, shifted from row 4)* | |

> Two rows inserted → the footer shifts from row 4 to row 6. openpyxl shifts all cell references, formatting, and merges automatically.

---

## 4. Outer join — more template rows than data

Template rows beyond the DataFrame length are **not deleted** (use `inner` if you want them cleared). They are left unmatched — their data columns keep whatever was in the template.

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | `{{ sales \| table(join=outer) }}` |
| 3   | South  | |
| 4   | East   | |
| 5   | West   | |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South"],
    "Revenue": [1000, 800],
})
```

**Output**

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | *(left as-is)* |
| 5   | West   | *(left as-is)* |

---

## 5. Right join

Overwrites template rows **top-down in DataFrame order** — no key matching. If the DataFrame is longer, rows are inserted. If shorter, extra template rows are cleared.

**Tag**
```
{{ sales | table(join=right) }}
```

**Template**
*(3 placeholder rows)*

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | *placeholder* | `{{ sales \| table(join=right) }}` |
| 3   | *placeholder* | |
| 4   | *placeholder* | |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East", "West"],
    "Revenue": [1000, 800, 600, 400],
})
```

**Output** *(4 rows, 1 inserted)*

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  |
| 5   | West   | 400  | ← inserted

> Template row keys are irrelevant — rows are written positionally in DataFrame order.

**Right join with fewer rows (clearing):**

Same template (3 rows), DataFrame with only 2 rows:

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | *(cleared)* | *(cleared)* |

---

## 6. Boundary markers: `end_table`

By default, `_find_last_data_row` auto-detects the bottom of the table zone by scanning for an empty join column. Use `{{ end_table }}` on its **own row** (no join value) to set an explicit boundary. The `end_table` row is deleted from the output.

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | `{{ sales \| table(join=outer) }}` |
| 3   | South  | |
| 4   | *(empty)* | `{{ end_table }}` |
| 5   | Grand Total | 9999 |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East"],
    "Revenue": [1000, 800, 600],
})
```

**Output**

| Row | Region | Revenue |
|-----|--------|---------|
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  | ← inserted (end_table row was here, now deleted)
| 5   | Grand Total | 9999 | ← shifted from row 5 (no change, the end_table row was on row 4, then deleted after insertion)

> The `end_table` row is deleted **after** row insertion, so the Grand Total row ends up at row 5.

**`end_table` on a data row (Option B):**

```
| 3   | South  | {{ end_table }} |
```

When `end_table` is on a row that **has a join value**, the marker tag is cleared but the row is kept as the last data row. Extras go after it.

---

## 7. Boundary markers: `end_table | insert=above` (pin a row)

This is **Option C** — the `end_table` row is part of the table zone AND gets matched by join key, but any extra rows are inserted **before** it. Use this to pin a row (like a Total) at the bottom.

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue *(headers)* |
| 2   | North  | `{{ sales \| table(join=outer) }}` |
| 3   | South  | |
| 4   | Total  | `{{ end_table \| insert=above }}` |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East", "West", "Total"],
    "Revenue": [1000, 800, 600, 400, 2800],
})
```

**Output** *(2 extra rows: East, West)*

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue |
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  | ← inserted before Total
| 5   | West   | 400  | ← inserted before Total
| 6   | Total  | 2800 | ← pinned, shifted down

> Total is matched by its join key value and filled from the DataFrame. It always stays last.

---

## 8. Boundary markers: `insert_data`

`{{ insert_data }}` marks the **exact insertion point** for outer join extras. Its row is deleted before processing. Useful when you want extra rows to appear at a specific position inside the template zone rather than at the bottom.

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue |
| 2   | North  | `{{ sales \| table(join=outer) }}` |
| 3   | South  | |
| 4   | *(empty)* | `{{ insert_data }}` |
| 5   | Total  | 9999 |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "South", "East", "West"],
    "Revenue": [1000, 800, 600, 400],
})
```

**Output**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue |
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  | ← inserted (insert_data row was here)
| 5   | West   | 400  | ← inserted
| 6   | Total  | 9999 | ← shifted from row 5

> Unlike `end_table`, the `insert_data` row carries no join value — it's a pure insertion-point marker. Total's value is **not** matched from the DataFrame; it remains as hardcoded template content.

---

## 9. Pinned Total row with `placeholder=true`

When the template has **zero** real data rows between the headers and the pinned row, you need a placeholder row as a style source. `placeholder=true` makes the writer use that row's style for inserted rows, then **delete** it if its join column is blank (unmatched).

Combine with `{{ end_table | insert=above }}` to pin the Total row.

**Template**

| Row | Index | Value |
|-----|-------|-------|
| 1   | Index | Value *(headers)* |
| 2   | *(blank)* | `{{ data \| table(join=outer, placeholder=true) }}` |
| 3   | Total | `{{ end_table \| insert=above }}` |

> Row 2's join column (Index) is blank → it will be deleted from the output.
> Row 2's **style** is still used as the template for inserted rows.

**DataFrame**
```python
pl.DataFrame({
    "Index": ["a", "b", "c", "Total"],
    "Value": [10, 20, 30, 60],
})
```

**Output**

| Row | Index | Value |
|-----|-------|-------|
| 1   | Index | Value |
| 2   | a     | 10   | ← placeholder row deleted; data starts here
| 3   | b     | 20   |
| 4   | c     | 30   |
| 5   | Total | 60   | ← pinned, matched by key

> If `placeholder=true` is omitted, row 2 stays blank in the output, resulting in a phantom empty row before row `a`.

---

## 10. Sorted outer join (`order_by`)

Outer join only. Sorts the **upper zone** (all rows above the `end_table | insert=above` boundary, if present) by a column. Lower zone rows (pinned via `insert=above`) are written in template order, separately.

**Tags**
```
{{ data | table(join=outer, order_by=asc) }}           ← sort join column ascending
{{ data | table(join=outer, order_by=desc) }}          ← sort join column descending
{{ data | table(join=outer, order_by=Revenue:desc) }}  ← sort by named column (desc)
```

**Template**

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue |
| 2   | East   | `{{ data \| table(join=outer, order_by=Revenue:desc) }}` |
| 3   | North  | |
| 4   | Total  | `{{ end_table \| insert=above }}` |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["North", "East", "South", "West", "Total"],
    "Revenue": [1000, 600, 800, 400, 2800],
})
```

**Output** *(upper zone sorted by Revenue descending; 2 rows inserted)*

| Row | Region | Revenue |
|-----|--------|---------|
| 1   | Region | Revenue |
| 2   | North  | 1000 |
| 3   | South  | 800  |
| 4   | East   | 600  |
| 5   | West   | 400  | ← inserted
| 6   | Total  | 2800 | ← pinned, shifted

> Template rows with no matching join key are included in the sort output with `None` data columns (sorted to the end after non-null rows).

---

## 11. Fill missing values (`fill=`)

Replaces `None` (missing) values in filled cells with a fallback, either globally or per column.

**Tags**
```
{{ data | table(join=outer, fill=0) }}               ← replace all None with 0
{{ data | table(join=outer, fill=Revenue:0;Margin:N/A) }}  ← per-column
```

**DataFrame** *(with missing values)*
```python
pl.DataFrame({
    "Region":  ["North", "South", "East"],
    "Revenue": [1000, None, 600],
    "Margin":  [0.20, 0.15, None],
})
```

**Output with `fill=0`**

| Region | Revenue | Margin |
|--------|---------|--------|
| North  | 1000 | 0.20 |
| South  | 0    | 0.15 |
| East   | 600  | 0    |

**Output with `fill=Revenue:0;Margin:N/A`**

| Region | Revenue | Margin |
|--------|---------|--------|
| North  | 1000 | 0.20 |
| South  | 0    | 0.15 |
| East   | 600  | N/A  |

> Fill also applies to template rows that are **unmatched** in a left/outer join — fills in the pre-existing template cell values through the same spec.

---

## 12. Style source for inserted rows (`style=first|last`)

When extra rows are inserted, their formatting is copied from a template row. By default (`style=last`) this is `last_tmpl_row`. Use `style=first` to copy from the tag row instead.

**Use case:** Template has uniform data rows plus a styled Total row at the bottom. You want inserted rows to look like data rows, not like the Total row.

**Template**

| Row | Index | Value | *Styling* |
|-----|-------|-------|-----------|
| 1   | Index | Value | headers |
| 2   | a     | `{{ data \| table(join=outer, style=first) }}` | plain |
| 3   | Total | 100   | **bold, yellow fill** |

**DataFrame**
```python
pl.DataFrame({
    "Index": ["a", "b", "c"],
    "Value": [10, 20, 30],
})
```

**Output**

| Row | Index | Value | *Styling* |
|-----|-------|-------|-----------|
| 1   | Index | Value | headers |
| 2   | a     | 10 | plain |
| 3   | b     | 20 | plain ← copied from row 2 (style=first) |
| 4   | c     | 30 | plain ← copied from row 2 (style=first) |
| 5   | Total | 100 | **bold, yellow** ← shifted |

**With `style=last` (default):**

| Row | Index | Value | *Styling* |
|-----|-------|-------|-----------|
| 3   | b     | 20 | **bold, yellow** ← copied from Total row |
| 4   | c     | 30 | **bold, yellow** ← copied from Total row |

> `style` only controls inserted rows. Template rows that are matched and filled keep their own existing formatting.

---

## 13. Positional fill (`positional=true`)

Writes the DataFrame directly by position — no header matching, no join column. The tag cell is the top-left of the write region. All cells are overwritten in order.

**Tag**
```
{{ data | table(positional=true) }}
```

**Template**

```
         A         B         C
  1   [ label ]  
  2              [ data tag ]
```

Cell B2 contains `{{ data | table(positional=true) }}`

**DataFrame**
```python
pl.DataFrame({
    "col1": [1, 2, 3],
    "col2": [4, 5, 6],
})
```

**Output** *(tag cell is top-left of write region)*

```
         A       B    C
  1   [ label ]  
  2              1    4
  3              2    5
  4              3    6
```

> No headers are written — only the data values. The column names in the DataFrame are ignored.

---

## 14. Custom join column (`on=`)

By default the join key column is the column to the **left** of the tag cell, matched against the DataFrame column whose header appears in the row above that cell. Use `on=` to override the DataFrame column name to join against.

**Use case:** Template header says "Sector" but DataFrame column is named "sector_name".

**Tag**
```
{{ data | table(join=outer, on=sector_name) }}
```

**Template**

| Sector | Revenue |
|--------|---------|
| Energy | `{{ data \| table(join=outer, on=sector_name) }}` |
| Tech   | |

**DataFrame**
```python
pl.DataFrame({
    "sector_name": ["Energy", "Tech", "Finance"],
    "Revenue":     [500, 300, 200],
})
```

**Output**

| Sector | Revenue |
|--------|---------|
| Energy | 500 |
| Tech   | 300 |
| Finance | 200 | ← inserted

---

## 15. Combining options

Options compose freely on the same tag.

### Pinned Total + sorted data + plain-row style for inserts

```
{{ data | table(join=outer, order_by=Revenue:desc, style=first) }}
```
with `{{ end_table | insert=above }}` on the Total row.

**Template**

| Row | Region | Revenue | *Style* |
|-----|--------|---------|---------|
| 1   | Region | Revenue | headers |
| 2   | *(placeholder)* | `{{ data \| table(join=outer, order_by=Revenue:desc, placeholder=true, style=first) }}` | plain |
| 3   | Total  | `{{ end_table \| insert=above }}` | bold |

**DataFrame**
```python
pl.DataFrame({
    "Region":  ["East", "North", "South", "Total"],
    "Revenue": [600, 1000, 800, 2400],
})
```

**Output** *(sorted desc by Revenue; placeholder row deleted; Total pinned)*

| Row | Region  | Revenue | *Style* |
|-----|---------|---------|---------|
| 1   | Region  | Revenue | headers |
| 2   | North   | 1000    | plain |
| 3   | South   | 800     | plain ← inserted, style from row 2 |
| 4   | East    | 600     | plain ← inserted, style from row 2 |
| 5   | Total   | 2400    | bold (pinned, matched by key) |

### Outer join + fill + per-column fill

```
{{ data | table(join=outer, fill=Revenue:0;Margin:N/A) }}
```

Fills `None` Revenue with `0` and `None` Margin with `"N/A"` across all rows — both matched and unmatched template rows.

---

## Option Reference Summary

| Option | Values | Default | Applies to |
|--------|--------|---------|------------|
| `join=` | `left` `inner` `outer` `right` | `left` | all |
| `on=` | column name string | inferred from header | all |
| `order_by=` | `asc` `desc` `ColName` `ColName:asc` `ColName:desc` | *(none)* | `outer` only |
| `fill=` | scalar value or `col1:v1;col2:v2` | *(no fill)* | all |
| `positional=` | `true` | `false` | all |
| `placeholder=` | `true` | `false` | `outer` only |
| `style=` | `first` `last` | `last` | `outer` `right` |

| Structural marker | When to use |
|---|---|
| `{{ end_table }}` (own row) | Explicit table boundary; marker row deleted |
| `{{ end_table }}` (data row) | Explicit last data row; extras go after |
| `{{ end_table \| insert=above }}` (data row) | Pin a row last; extras inserted before it |
| `{{ insert_data }}` | Mark exact insertion point; row deleted |
