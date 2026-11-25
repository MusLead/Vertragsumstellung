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


def comparison_pairs(groups, content_keys):
    if not {"kundennummer", "kunde"} <= set(content_keys):
        return {}

    customer_indexes = [content_keys.index(key) for key in ("kundennummer", "kunde")]
    candidates = [{}, {}]
    for index, (signature, (first_rows, second_rows)) in enumerate(groups):
        if len(first_rows) == len(second_rows):
            continue
        customer = tuple(signature[i] for i in customer_indexes)
        if any(value in (None, "") for _, value in customer):
            continue
        side = 0 if len(first_rows) > len(second_rows) else 1
        candidates[side][index] = customer

    pairs = {}
    while candidates[0] and candidates[1]:
        best = [{}, {}]
        for side in (0, 1):
            for index, customer in candidates[side].items():
                distances = [
                    (sum(a != b for a, b in zip(groups[index][0], groups[other][0])), other)
                    for other, other_customer in candidates[1 - side].items()
                    if customer == other_customer
                ]
                if distances:
                    minimum = min(distance for distance, _ in distances)
                    nearest = [other for distance, other in distances if distance == minimum]
                    if len(nearest) == 1:
                        best[side][index] = nearest[0]
        found = [
            (first, second)
            for first, second in best[0].items()
            if best[1].get(second) == first
        ]
        if not found:
            break
        for first, second in found:
            pairs[first] = second
            del candidates[0][first]
            del candidates[1][second]
    return pairs


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
        report(f"Erste Datei: {first}")
        report(f"Zweite Datei: {second}")
        report(f"Tabellenblatt: {sheet or 'erstes Tabellenblatt'}")
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

                grouped_rows = list(groups.items())
                pairs = comparison_pairs(grouped_rows, content_keys)
                paired = set(pairs) | set(pairs.values())
                for index, (_, (first_rows, second_rows)) in enumerate(grouped_rows):
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
                        f"Erste Datei: {first_count} | Zweite Datei: {second_count} | "
                        f"Abweichung: {difference}"
                    )
                    for label, rows in (("Erste Datei", first_rows),
                                        ("Zweite Datei", second_rows)):
                        if rows:
                            report(
                                f"  Zeilen {label}: "
                                + ", ".join(str(number) for number, _ in rows)
                            )
                    row = (first_rows or second_rows)[0][1]
                    for key in ("kundennummer", "kunde"):
                        if key in row and row[key] not in (None, ""):
                            report(f"  {names[key]}: {row[key]!r}")
                    if index not in paired:
                        report("  Kein eindeutiger Vergleichspartner für die überschüssigen Zeilen.")

                if pairs:
                    report("Mögliche Änderungen (eindeutig ähnlichste Zeilengruppen):")
                    for first_index, second_index in pairs.items():
                        first_signature, (first_rows, first_matches) = grouped_rows[first_index]
                        second_signature, (second_matches, second_rows) = grouped_rows[second_index]
                        report(
                            "  Erste Datei Zeilen: "
                            + ", ".join(str(number) for number, _ in first_rows[len(first_matches):])
                            + " | Zweite Datei Zeilen: "
                            + ", ".join(str(number) for number, _ in second_rows[len(second_matches):])
                        )
                        for key, before_value, after_value in zip(
                            content_keys, first_signature, second_signature
                        ):
                            if before_value != after_value:
                                report(f"    {names[key]}: {before_value[1]!r} -> {after_value[1]!r}")
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
