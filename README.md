# Parts List Comparison

Compares two parts lists in an Excel workbook and writes the differences to a
`Comparison` tab as plain values (no formulas).

## Files

| File | Purpose |
|------|---------|
| `compare.py` | The comparison script. |
| `parts_comparison_template.xlsx` | Blank workbook with `List 1`, `List 2`, `Comparison` tabs. Paste your data here. |
| `sample_parts.xlsx` | Sample workbook with test data and the generated `Comparison` tab. |
| `make_sample.py` | Regenerates the template and sample workbooks. |
| `test_compare.py` | Regression test that checks the sample output row by row. |

## Setup

Requires Python 3.8+.

```
pip install openpyxl
```

## Run

1. Open `parts_comparison_template.xlsx` (or any workbook with the same layout)
   and paste your raw data into `List 1` and `List 2`. Header in row 1:

   | A | B | C | D |
   |---|---|---|---|
   | Item # | Quantity | Part # | Part Description |

2. Save and close the workbook, then run:

   ```
   python compare.py "path/to/workbook.xlsx"
   ```

3. Reopen the workbook. The `Comparison` tab is rewritten every run.

## How rows are matched

- Match key is Item # + Part # + Part Description. All three must match after
  trimming whitespace, ignoring case.
- Duplicate keys within a list stay as separate rows. They are paired in order
  of appearance (1st with 1st, 2nd with 2nd). Leftovers show as missing from
  the other list.
- A blank or non-numeric Quantity counts as 0.
- Fully blank rows are ignored.

## Output (Comparison tab)

`Item # | Part # | Part Description | Qty List 1 | Qty List 2 | Difference (L2 - L1) | Status`

Sorted by Item # then Part # (numerically when the values are numbers).
Header is bold with a filter, the top row is frozen, and columns are autofit.

| Status | Meaning | Row color |
|--------|---------|-----------|
| Match | In both lists, same quantity | green |
| Qty Different | In both lists, different quantity | yellow |
| Only in List 1 | No matching row in List 2 | red |
| Only in List 2 | No matching row in List 1 | blue |

A summary block in columns I:J counts each status plus the total.

## Test

```
python test_compare.py
```

Builds the sample data in a temp file, runs the comparison twice (to confirm
old results are cleared), and checks every output row, color, and the summary.
