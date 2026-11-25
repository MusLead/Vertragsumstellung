import os
import shutil

from logging_config import setup_logger

logger, LOG_FILE = setup_logger(__file__)

import pandas as pd
from project_config import CONFIG_FILE, get_configured_path, load_project_config

# ============================================================
# Files
# ============================================================

CONTRACT_FILE = get_configured_path("contract_workbook")
SAP_FILE = get_configured_path("sap_export")
OUTPUT_FILE = get_configured_path("enriched_workbook")
SAP_LOG_COLUMN = "SAP-LOG"


def normalize_sap_log_column(sap):
    log_headers = [
        column
        for column in sap.columns
        if "log" in column.casefold() or "error" in column.casefold()
    ]
    if len(log_headers) > 1:
        raise ValueError(
            "SAP export has multiple error/log columns: "
            + ", ".join(log_headers)
        )

    if log_headers:
        sap = sap.rename(columns={log_headers[0]: SAP_LOG_COLUMN})
        logger.info(
            "Using SAP export column %s as the error/log field",
            log_headers[0],
        )
        return sap

    sap[SAP_LOG_COLUMN] = ""
    logger.info("SAP export has no error/log column; using blank values")
    return sap


def load_customer_name_rules():
    config = load_project_config()

    if not isinstance(config, dict) or not isinstance(config.get("rules"), list):
        raise ValueError(
            f"{CONFIG_FILE} must contain a top-level 'rules' list"
        )

    rules = []
    for rule_number, rule in enumerate(config["rules"], start=1):
        if not isinstance(rule, dict):
            raise ValueError(f"Rule {rule_number} in {CONFIG_FILE} must be a mapping")

        unknown_keys = set(rule) - {
            "Kundennummer",
            "Kunde",
            "set",
            "overrideCell",
        }
        if unknown_keys:
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} has unsupported keys: "
                + ", ".join(sorted(unknown_keys))
            )

        customer_number = rule.get("Kundennummer")
        customer_name = rule.get("Kunde")
        changes = rule.get("set")
        override_cell = rule.get("overrideCell", False)

        if not isinstance(customer_number, str) or not customer_number.strip():
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} must have a quoted, "
                "non-empty Kundennummer"
            )
        if not isinstance(customer_name, str) or not customer_name.strip():
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} must have a non-empty Kunde"
            )
        if not isinstance(changes, dict) or not changes:
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} must have a non-empty 'set' mapping"
            )
        if not isinstance(override_cell, bool):
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} must set overrideCell to yes or no"
            )
        invalid_columns = set(changes) - {"Name1", "Name2"}
        if invalid_columns:
            raise ValueError(
                f"Rule {rule_number} in {CONFIG_FILE} has unsupported columns: "
                + ", ".join(sorted(invalid_columns))
            )
        normalized_changes = {}
        for column, value in changes.items():
            if not isinstance(value, str) or not value.strip():
                raise ValueError(
                    f"Rule {rule_number} in {CONFIG_FILE} must give {column} "
                    "a non-empty text value"
                )
            normalized_changes[column] = value.strip()

        rules.append(
            {
                "Kundennummer": customer_number.strip(),
                "Kunde": customer_name.strip(),
                "set": normalized_changes,
                "overrideCell": override_cell,
                "rule_number": rule_number,
            }
        )

    return rules


def apply_customer_name_rules(merged, rules):
    for rule in rules:
        matching_rows = (
            merged["Kundennummer"].astype(str).str.strip().eq(rule["Kundennummer"])
            & merged["Kunde"].astype(str).str.strip().eq(rule["Kunde"])
        )
        if not matching_rows.any():
            raise ValueError(
                f"Rule {rule['rule_number']} in {CONFIG_FILE} did not match "
                f"Kundennummer {rule['Kundennummer']} and Kunde {rule['Kunde']!r}"
            )

        for column, target_value in rule["set"].items():
            if column not in merged.columns:
                raise ValueError(
                    f"Rule {rule['rule_number']} in {CONFIG_FILE} targets missing "
                    f"column {column!r}"
                )

            current_values = (
                merged.loc[matching_rows, column]
                .fillna("")
                .astype(str)
                .str.strip()
            )
            empty_rows = current_values.index[current_values.eq("")]
            mismatch_rows = current_values.index[
                current_values.ne("") & current_values.ne(target_value)
            ]

            merged.loc[empty_rows, column] = target_value
            if not len(mismatch_rows):
                logger.info(
                    "Applied customer-name rule %s to %s empty %s cell(s)",
                    rule["rule_number"],
                    len(empty_rows),
                    column,
                )
                continue

            should_override = rule["overrideCell"]
            if not should_override:
                existing_values = ", ".join(
                    sorted(set(current_values.loc[mismatch_rows].tolist()))
                )
                try:
                    answer = input(
                        f"Rule {rule['rule_number']}: {column} for Kundennummer "
                        f"{rule['Kundennummer']} / Kunde {rule['Kunde']!r} is "
                        f"already filled in {len(mismatch_rows)} matching row(s) "
                        f"with {existing_values!r}. Replace it with "
                        f"{target_value!r}? [y/N]: "
                    )
                except EOFError as error:
                    raise RuntimeError(
                        f"Rule {rule['rule_number']} requires confirmation. "
                        "Set overrideCell: yes to allow the overwrite without a prompt."
                    ) from error
                should_override = answer.strip().casefold() in {
                    "y",
                    "yes",
                    "j",
                    "ja",
                }

            if should_override:
                merged.loc[mismatch_rows, column] = target_value
                logger.info(
                    "Overrode %s populated %s cell(s) using customer-name rule %s",
                    len(mismatch_rows),
                    column,
                    rule["rule_number"],
                )
            else:
                logger.info(
                    "Kept existing %s values for customer-name rule %s",
                    column,
                    rule["rule_number"],
                )


# ============================================================
# Read Files
# ============================================================

logger.info("Loading contract file: %s", CONTRACT_FILE)
try:
    contracts = pd.read_excel(
        CONTRACT_FILE,
        sheet_name="uniform",
        dtype=str
    )
except PermissionError:
    logger.exception("Could not open contract workbook %s", CONTRACT_FILE)
    raise SystemExit(
        f"Could not open {CONTRACT_FILE}. Please close the Excel workbook "
        "if it is open, then run this script again."
    ) from None
logger.info("Loaded %s contract rows", len(contracts))

logger.info("Loading SAP export: %s", SAP_FILE)
sap = pd.read_csv(
    SAP_FILE,
    sep=";",
    dtype=str,
    encoding="utf-8"
)
logger.info("Loaded %s SAP rows", len(sap))

# ============================================================
# Cleanup
# ============================================================

contracts.columns = contracts.columns.str.strip()
sap.columns = sap.columns.str.strip()
sap = normalize_sap_log_column(sap)

contracts = contracts.fillna("")
sap = sap.fillna("")

contracts["Kundennummer"] = (
    contracts["Kundennummer"]
    .astype(str)
    .str.strip()
)

intentional_skip_rows = contracts["Kundennummer"].eq("-")
intentional_skip_count = int(intentional_skip_rows.sum())
if intentional_skip_rows.any():
    logger.info(
        "Ignoring %s contract row(s) with Kundennummer '-' as intentional skips",
        intentional_skip_count,
    )
    contracts = contracts.loc[~intentional_skip_rows].copy()

sap["Kundennummer"] = (
    sap["Kundennummer"]
    .astype(str)
    .str.strip()
)

contract_names = contracts.loc[
    contracts["Kundennummer"].ne("")
    & contracts["Kunde"].astype(str).str.strip().ne(""),
    ["Kundennummer", "Kunde"],
].copy()
contract_names["Kunde"] = contract_names["Kunde"].astype(str).str.strip()
contract_names = contract_names.drop_duplicates()

customer_name_counts = contract_names.groupby("Kundennummer")["Kunde"].nunique()
conflicting_customer_ids = customer_name_counts[
    customer_name_counts > 1
].index
customer_name_conflicts = (
    contract_names[contract_names["Kundennummer"].isin(conflicting_customer_ids)]
    .groupby("Kundennummer")["Kunde"]
    .agg(list)
)

for customer_number, customer_names in customer_name_conflicts.items():
    logger.warning(
        "Kundennummer %s is associated with multiple Kunde names: %s",
        customer_number,
        " / ".join(customer_names),
    )

# ============================================================
# Remove duplicate SAP customers
# (keep first occurrence)
# ============================================================

sap_unique = sap.drop_duplicates(
    subset=["Kundennummer"],
    keep="first"
)

# ============================================================
# LEFT JOIN
# ============================================================

merged = contracts.merge(
    sap_unique,
    how="left",
    on="Kundennummer"
)
apply_customer_name_rules(merged, load_customer_name_rules())

# ============================================================
# Logging
# ============================================================

total_contract_rows = len(contracts)
total_merged_rows = len(merged)

missing_customer_data = merged[
    merged["Name1"].fillna("").eq("")
]

summary_file = LOG_FILE.with_name(f"{LOG_FILE.stem}_summary.txt")
with open(summary_file, "w", encoding="utf-8") as f:
    f.write("=== LEFT JOIN LOG ===\n\n")
    f.write(f"Contract rows      : {total_contract_rows}\n")
    f.write(f"Merged rows        : {total_merged_rows}\n")
    f.write(f"Missing customers  : {len(missing_customer_data)}\n\n")
    f.write(f"Intentional '-' skips: {intentional_skip_count}\n")
    f.write(f"IDs with multiple Kunde names: {len(customer_name_conflicts)}\n\n")

    if not customer_name_conflicts.empty:
        f.write("Kundennummer values associated with multiple Kunde names:\n\n")
        for customer_number, customer_names in customer_name_conflicts.items():
            f.write(f"{customer_number}: {' / '.join(customer_names)}\n")
        f.write("\n")

    logger.info(
        "Join complete: %s contract rows, %s merged rows, %s missing customers",
        total_contract_rows,
        total_merged_rows,
        len(missing_customer_data),
    )

    if len(missing_customer_data) > 0:
        f.write("Customers not found in SAP:\n\n")

        for knr in (
            missing_customer_data["Kundennummer"]
            .dropna()
            .unique()
        ):
            f.write(f"{knr}\n")
            logger.warning("Customer not found in SAP: %s", knr)

# ============================================================
# Save
# ============================================================

if os.path.exists(OUTPUT_FILE):
    logger.info("Removing existing enriched workbook: %s", OUTPUT_FILE)
    try:
        os.remove(OUTPUT_FILE)
    except PermissionError:
        logger.exception("Could not remove existing workbook %s", OUTPUT_FILE)
        raise SystemExit(
            f"Could not replace {OUTPUT_FILE}. Please close it if it is open, "
            "then run this script again."
        ) from None

shutil.copyfile(CONTRACT_FILE, OUTPUT_FILE)
with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    mode="a",
    if_sheet_exists="replace",
) as writer:
    merged.to_excel(
        writer,
        index=False,
        sheet_name="uniform",
    )
logger.info("Enriched workbook saved: %s", OUTPUT_FILE)

print()
print("Done.")
print(f"Rows before join : {total_contract_rows}")
print(f"Rows after join  : {total_merged_rows}")
print(f"Output           : {OUTPUT_FILE}")
print(f"Log              : {LOG_FILE}")
print(f"Summary          : {summary_file}")
logger.info("Run completed successfully")