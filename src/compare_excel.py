"""Compare cached Excel cell values by content, matching columns by header."""
import argparse
from contextlib import ExitStack, closing
from datetime import datetime

from openpyxl import load_workbook

from project_config import get_configured_path


def normal(value):
    return str(value).strip().casefold()


def is_acknowledged(value):
    return normal(value) == "ok"


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
    missing = {"anmerkungen"} - names.keys()
    if missing:
        raise ValueError(f"{ws.title}: Missing headers: {', '.join(sorted(missing))}")

    def records():
        for number, row in enumerate(rows, 2):
            if any(value is not None for value in row[len(names):]):
                raise ValueError(f"{ws.title}, row {number}: Data under a missing header")
            yield dict(zip(names, row)) if any(v is not None for v in row) else {}

    return names, records()


def compare(first, second, sheet="uniform"):
    counts = dict(equal=0, different=0, warnings=0)
    logs_folder = get_configured_path("logs_folder")
    logs_folder.mkdir(parents=True, exist_ok=True)
    started_at = datetime.now()
    timestamp = started_at.strftime("%Y%m%d_%H%M%S_%f")
    log_file = logs_folder / f"compare_excel_{timestamp}.log"

    with log_file.open("w", encoding="utf-8") as log:
        def report(message):
            print(message)
            log.write(message + "\n")

        report(f"Started: {started_at:%Y-%m-%d %H:%M:%S}")
        report(f"Comparing {first} with {second} (sheet: {sheet or 'first sheet'})")
        try:
            with ExitStack() as stack:
                tables = []
                for path in (first, second):
                    book = stack.enter_context(
                        closing(load_workbook(path, read_only=True, data_only=True))
                    )
                    tables.append(read_table(book, sheet))
                (names, before), (other_names, after) = tables
                if names.keys() != other_names.keys():
                    raise ValueError(
                        f"Different columns: {sorted(names.keys() ^ other_names.keys())}"
                    )

                content_keys = [key for key in names if key != "anmerkungen"]
                groups = {}
                for file_index, rows in enumerate((before, after)):
                    for number, row in enumerate(rows, 2):
                        if not row:
                            continue
                        signature = tuple(
                            (isinstance(row[key], bool), row[key])
                            for key in content_keys
                        )
                        group = groups.setdefault(signature, ([], []))
                        group[file_index].append((number, row))

                for first_rows, second_rows in groups.values():
                    first_count, second_count = len(first_rows), len(second_rows)
                    counts["equal"] += min(first_count, second_count)
                    difference = abs(first_count - second_count)
                    if not difference:
                        continue

                    acknowledged = any(
                        is_acknowledged(row.get("anmerkungen"))
                        for _, row in first_rows + second_rows
                    )
                    warning = acknowledged or max(first_count, second_count) > 1
                    counts["different"] += difference
                    counts["warnings"] += difference if warning else 0
                    report(
                        f"{'WARNUNG' if warning else 'Unterschied'} | "
                        f"Datei 1: {first_count} | Datei 2: {second_count} | "
                        f"Abweichung: {difference}"
                    )
                    for file_number, rows in enumerate((first_rows, second_rows), 1):
                        if rows:
                            report(
                                f"  Zeilen Datei {file_number}: "
                                + ", ".join(str(number) for number, _ in rows)
                            )
                    row = (first_rows or second_rows)[0][1]
                    for key in content_keys:
                        report(f"  {names[key]}: {row[key]!r}")
        except (ValueError, KeyError, OSError) as exc:
            report(f"ERROR: {exc}")
            raise

        report(
            "Gleiche Zeilen: {equal} | Abweichende Zeilen: {different} | "
            "Davon Warnungen: {warnings}".format(**counts)
        )
        report(f"Logdatei: {log_file}")
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
    parser.add_argument(
        "--sheet",
        default="uniform",
        help="Worksheet name (default: uniform)",
    )
    args = parser.parse_args()
    try:
        compare(args.first, args.second, args.sheet)
    except (ValueError, KeyError, OSError) as exc:
        parser.exit(2, f"Error: {exc}\n")
