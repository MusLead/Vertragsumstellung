"""Compare two Excel tables by physical row, matching columns by header."""
import argparse
from contextlib import ExitStack, closing
from itertools import zip_longest

from openpyxl import load_workbook

from project_config import get_configured_path


def normal(value):
    return str(value).strip().casefold()


def read_table(book, sheet):
    ws = book[sheet] if sheet else book.worksheets[0]
    rows = ws.values
    headers = list(next(rows, ()))
    while headers and headers[-1] is None:
        headers.pop()
    names = {}
    for value in headers:
        key = normal(value)
        if key == "kundenummer":
            key = "kundennummer"
        if value is None or not key or key in names:
            raise ValueError(f"{ws.title}: Empty or duplicate header: {value!r}")
        names[key] = str(value)
    missing = {"kundennummer", "kunde", "status"} - names.keys()
    if missing:
        raise ValueError(f"{ws.title}: Missing headers: {', '.join(sorted(missing))}")

    def records():
        for number, row in enumerate(rows, 2):
            if any(value is not None for value in row[len(names):]):
                raise ValueError(f"{ws.title}, row {number}: Data under a missing header")
            yield dict(zip(names, row)) if any(v is not None for v in row) else {}

    return names, records()


def compare(first, second, sheet=None):
    counts = dict(equal=0, different=0, warnings=0)
    with ExitStack() as stack:
        tables = []
        for path in (first, second):
            book = stack.enter_context(closing(load_workbook(path, read_only=True)))
            tables.append(read_table(book, sheet))
        (names, before), (other_names, after) = tables
        if names.keys() != other_names.keys():
            raise ValueError(f"Different columns: {sorted(names.keys() ^ other_names.keys())}")

        for number, (a, b) in enumerate(zip_longest(before, after, fillvalue={}), 2):
            if not (a or b):
                continue
            changes = []
            if not a or not b:
                changes.append(f"Row missing or empty in file {1 if not a else 2}")
            else:
                for key, label in names.items():
                    x, y = a[key], b[key]
                    if key == "status" and normal(x) == normal(y) == "ok":
                        continue
                    if x != y or isinstance(x, bool) != isinstance(y, bool):
                        changes.append((key, f"{label}: {x!r} -> {y!r}"))
            if not changes:
                counts["equal"] += 1
                continue
            counts["different"] += 1
            ok = any(normal(row.get("status")) == "ok" for row in (a, b))
            warning = ok and (not a or not b or any(key != "status" for key, _ in changes))
            counts["warnings"] += int(warning)
            customer = a or b
            print(f"{'WARNUNG' if warning else 'Unterschied'} | Zeile {number} | "
                  f"Kundennummer {customer['kundennummer']} | Kunde {customer['kunde']}")
            for change in changes:
                print("  " + (change[1] if isinstance(change, tuple) else change))
    print("Gleiche Zeilen: {equal} | Abweichende Zeilen: {different} | "
          "Davon Warnungen: {warnings}".format(**counts))
    return counts


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "first",
        nargs="?",
        default=str(get_configured_path("comparison_first")),
    )
    parser.add_argument(
        "second",
        nargs="?",
        default=str(get_configured_path("comparison_second")),
    )
    parser.add_argument("--sheet", help="Worksheet name (default: first sheet in each file)")
    args = parser.parse_args()
    try:
        compare(args.first, args.second, args.sheet)
    except (ValueError, KeyError, OSError) as exc:
        parser.exit(2, f"Error: {exc}\n")
