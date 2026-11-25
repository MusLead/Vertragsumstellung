# Journal

## 2026-11-14 — Flag customer numbers linked to different names

Added a check in `join_python.py` to detect a `Kundennummer` associated with
multiple distinct, nonblank `Kunde` names. Names are trimmed before comparing,
so extra spaces do not create a false conflict. Each conflict is logged as a
`WARNING`, and the join summary lists the affected customer numbers and names.
This flags records for review; it does not establish whether the names refer to
different legal customers.

### Excel check

In the `uniform` sheet, the unedited enriched workbook uses column A for
`Kundennummer` and column B for `Kunde`. Add this formula in a helper column on
row 2 and fill it down. It flags a row when the same customer number appears
with a different nonblank name:

```excel
=IF(AND(A2<>"";B2<>"";COUNTIFS($A$2:$A$10000;A2;$B$2:$B$10000;"<>";$B$2:$B$10000;"<>"&B2)>0);"CHECK";"")
```

Increase `10000` if the data extends beyond row 10000. If Excel uses commas as
formula separators, replace the semicolons with commas.
