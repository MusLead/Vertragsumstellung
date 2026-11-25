# EVB-IT Contract Automation

This project automates the end-to-end EVB-IT contract creation process by combining SAP customer master data retrieval, data enrichment, Microsoft Office automation, and document generation.

The solution was developed as part of the AIS Automation Initiative to reduce manual workload, improve data quality, and establish a scalable and repeatable contract generation process within B. Braun.

---

## Overview

The automation workflow consists of four main stages:

1. Maintain contractual data in Microsoft Excel
2. Retrieve customer master data from SAP
3. Enrich contract data using a LEFT JOIN process
4. Generate and format EVB-IT contract documents automatically

This approach ensures that customer information is always obtained directly from SAP while preserving the original ordering structure of each customer.

## Project layout and local configuration

- Python scripts and shared helpers are in `src/`.
- Excel workbooks, SAP exports, and the Word contract template are in `res/`.
- Project notes are in `docs/`; this `README.md` remains in the project root.
- Runtime input, output, template, comparison, and log paths are read from the
  root `config.yaml`.

The YAML configuration and local business data are intentionally ignored by
Git. Keep a local `config.yaml` with a top-level `files` mapping containing
`contract_workbook`, `sap_export`, `enriched_workbook`, `contract_template`,
`contracts_output`, `logs_folder`, `comparison_first`, and `comparison_second`,
as well as the optional customer-specific `rules` list. All configured relative
paths are resolved from the project root. Do not commit the configuration,
workbooks, SAP exports, generated contracts, or archived data.

### Quick reference (Auto Fill In The Blank)

Video overview: https://youtube.com/shorts/JHaMjYHJkuo?is=Mq-agDmtWtR14Y5B

> Note: Before using the Excel contract workbook, make sure all cells are unmerged. Merged cells can break the data structure, cause false information, import issues, and interfere with the join and generation logic.

STEPS:

1. Select all the cells that needs to be filled in
2. `Ctr. + G` and then click button `special...`
3. Afterwards check the `blanks`
4. afterwards no need to click anything just type e.g. `=A2` so that it will take the value above e.g. in this example it will take the A2 as reference value. **THEN CLICK `Ctr. + Enter`**

---

## Contract Data Maintenance

The process starts with the contract overview workbook:

`res/EVB_IT-Vertragsübersicht_20260825.xlsx`

This workbook contains the business-relevant contract information, including:

IMPORTANT: The source Excel file must be cleaned so that all cells are unmerged before the automation runs. Merged cells can lead to incorrect row handling and downstream generation errors.

- Kundennummer
- Application / Product
- Menge (Quantity)
- Status
- Contract-specific attributes

A customer may appear multiple times because several applications can be assigned to the same customer.

Example:

| Kundennummer | Application | Menge |
|-------------|------------|--------|
| [redacted] | Monitor | 18 |
| [redacted] | DBMS | 5 |

These records represent the actual order structure and must remain unchanged throughout the entire automation process.

---

## Excel Data Formatting and Review

Use this process to standardize and validate the contract data before generating contracts. Keep the original workbook unchanged so that you can compare it with the standardized data.

1. **Unmerge cells.** Unmerge all cells in the workbook before processing it.
2. **Review `Vertragsinhalt`.** Check the `Vertragsinhalt` field for every record. If it is empty, review the record manually before removing it. Clarify unexpected or ambiguous entries rather than deleting them automatically.
3. **Compare the workbooks.** Display the original workbook and the workbook being standardized side by side. Do a high-level manual check that the customer names and `Vertragsinhalt` values are correct.
4. **Map the responsible teams.** Use the customer number to map each customer to the responsible Außendienst. Then use the Außendienst mapping to identify the responsible Innendienst.
5. **Verify against the original.** Check the standardized data against the original workbook. Do not change entries unless a correction has been confirmed.
6. **Create the `uniform` worksheet.** Once the data has been checked against the original, create a worksheet named `uniform` for the standardized data.
7. **Extract contract fields.** Use Excel formulas to extract the relevant information from `Vertragsinhalt` into new columns, including `Type`, `Menge`, and `Status`.
8. **Add pricing and name fields.** Add the `Special Price`, `Rabatt`, and `Mul. Name` columns.
9. **Apply the teams' feedback.** Update the data and `Status` based on the teams' responses. Use these status values:
   - `CONFIRMED`: The responsible Innendienst has confirmed the data; no changes are needed.
   - `REVISED`: The contract content has been corrected or adjusted.
   - `NR & REVISED`: The customer number and related data have been corrected.
   - `NR`: The customer number has been corrected.
10. **Review errors and data anomalies.** Check all remaining `ERROR` values, missing or incorrect customer numbers, special prices and discounts, and multiple customer names. Multiple names or numbers should be investigated unless the team has confirmed that they are intentional.

   To create a unique list of customer-number/name pairs from the `uniform` worksheet, enter this formula in an empty area:

   ```excel
   =LET(kunde;uniform!B2:B10000;nr;uniform!A2:A10000;UNIQUE(FILTER(HSTACK(IF(nr="";"";nr);kunde);kunde<>"")))
   ```

   The formula returns the customer number in the first column and the customer name in the second. It excludes rows with a blank customer name and retains rows with a blank customer number.

   To flag customer names associated with more than one customer number, add a helper column to `uniform`. With customer numbers in column A and customer names in column B, enter this formula in row 2 and fill it down:

   ```excel
   =IF(OR($A2="";$B2="");"";IF(COUNTIFS($B$2:$B$10000;$B2;$A$2:$A$10000;"<>"&$A2;$A$2:$A$10000;"<>")>0;"CHECK: same name, different customer number";""))
   ```

   Review every row flagged by the formula and confirm the correct customer number. The formula ignores rows with a blank name or customer number.

   As an alternative, this dynamic-array formula is intended to spill all unique customer-number/name pairs for names associated with more than one customer number:

   ```excel
   =LET(kunde;uniform!B2:B10000;nr;uniform!A2:A10000;paar;UNIQUE(FILTER(HSTACK(nr;kunde);(kunde<>"")*(nr<>"")));namen;CHOOSECOLS(paar;2);FILTER(paar;COUNTIF(namen;namen)>1;"No conflicts"))
   ```

   If this formula returns `#VALUE!` in your Excel, use the helper-column formula above instead.

11. **Retrieve SAP customer details.** Run `python src\sap_client.py` to retrieve each customer's address and complete name from SAP.
12. **Request a team review.** Send the workbook to the relevant colleagues. Ask them to mark confirmed records as `OK`; for corrections, ask them to set the status to `REVISED` and add a note describing the required change.

   **WICHTIG: Mandatory review documentation and protection rule:** Every time a review is completed, it must be documented what was checked, what was changed, by whom, and when. The review note should be stored with the workbook (or in the corresponding review log) so that the status of each review is traceable. After the review is finished, the reviewed cell(s) must be protected immediately. **THE CELL NEEDS TO BE PROTECTED** so no accidental edits can overwrite the validated content after approval.

13. **Generate the contracts.** Run `python src\generate_contracts_separate.py` and review the log files.
14. **Check the output.** Compare the number of generated contracts with the number of distinct `Kunde` groups in `uniform`. Investigate any discrepancy against the SAP customer list and the generation summary.
---

## SAP Customer Master Data Retrieval

### Script

`src/sap_client.py`

The SAP retrieval module connects directly to SAP GUI through SAP GUI Scripting and automatically accesses transaction **VD03**.

The script reads unique customer numbers from the `uniform` worksheet and checks them against the existing `res/sap_customer_export.csv`. Customer numbers already in the CSV are left unchanged and are not retrieved again. Only missing customer numbers are queued for SAP retrieval and appended to the CSV, preserving the existing export. If no customer numbers are missing, SAP is not opened.

For each queued customer number, the automation extracts:

- Customer Name
- Additional Customer Name
- Street
- House Number
- Postal Code
- City

The retrieved customer data is exported to:

`res/sap_customer_export.csv`

Example:

| Kundennummer | Name1 | City |
|-------------|--------|------|
| [redacted] | Example Customer | Example City |

The objective of this step is to ensure that all customer information originates directly from SAP and does not require manual maintenance.

---

## Contract Data Enrichment

### Script

`src/join_python.py`

Once customer master data has been retrieved successfully, the contract overview and SAP export are merged together.

The merge is performed using a SQL-style **LEFT JOIN** based on:

`Kundennummer`

### Why a LEFT JOIN?

The contract overview is considered the primary data source.

No contract records may be removed or aggregated during the enrichment process.

#### Contract Data

| Kundennummer | Application | Menge |
|-------------|------------|--------|
| [redacted] | Monitor | 18 |
| [redacted] | DBMS | 5 |

#### SAP Customer Data

| Kundennummer | Name1 | City |
|-------------|--------|------|
| [redacted] | Example Customer | Example City |

#### Enriched Result

| Kundennummer | Application | Menge | Name1 | City |
|-------------|------------|--------|--------|------|
| [redacted] | Monitor | 18 | Example Customer | Example City |
| [redacted] | DBMS | 5 | Example Customer | Example City |

Customer master data is automatically added to every relevant contract position.

This ensures that:

- No contract records are lost
- No application records are aggregated
- Multiple applications per customer remain separated
- Contract generation receives complete customer information

The resulting dataset is exported to:

`res/EVB_IT_Vertragsuebersicht_Enriched.xlsx`

Additionally, a validation summary and a timestamped chronological run log are generated in the `logs/` folder:

`logs/join_python_<timestamp>_summary.txt`

The summary contains:

- Total contract records
- Total enriched records
- Missing customer information
- Validation results
- Join statistics

### Customer-Specific SAP Name Corrections

`src/join_python.py` loads `config.yaml` and applies its rules after joining the contract data with SAP. Each rule matches both `Kundennummer` and the exact `Kunde` value, then sets `Name1` and/or `Name2` on the matching output rows. This allows a correction to target one customer-name variant without changing other rows that share its customer number.

For example, add a rule under `rules` for the relevant customer number and exact `Kunde` value, then set `Name2` to the required text. Quote customer numbers in YAML to preserve leading zeros.

- If the target cell is empty, the configured value is written automatically.
- If the target cell already contains a different value and `overrideCell: yes`, that value is replaced without a prompt.
- If the target cell already contains a different value and `overrideCell` is `no` or omitted, the script asks whether to replace it. Answer `y`/`yes` (or `j`/`ja`) to replace it; any other response keeps the existing value.

The configuration file includes a commented example. Replace its placeholders with the actual customer number, exact `Kunde` value, and desired name before adding it to `rules`.

#### YAML format

`config.yaml` must contain a top-level `files` mapping and a `rules` list. The `files` mapping defines the paths used by every script; each `rules` item describes one optional customer-specific correction:

```yaml

# files:
#   contract_workbook: "beispiel.xlsx"
#   sap_export: "res/beispiel.csv"
#   enriched_workbook: "res/beispiel_Enriched.xlsx"
#   contract_template: "res/beispiel_template.docx"
#   contracts_output: "output"
#   logs_folder: "logs"
#   comparison_first: "res/beispiel.xlsx"
#   comparison_second: "res/beispiel.xlsx"


# Add a rule for each customer number and exact Kunde value that needs a
# customer-name correction after the SAP join. Quote customer numbers.

# Example: fill Name2 for a specific customer variant.
# Replace the example values with the exact values from your workbook/export,
# then add this rule under "rules".
#
# rules:
#   - Kundennummer: "REPLACE_WITH_CUSTOMER_NUMBER"
#     Kunde: "REPLACE_WITH_EXACT_KUNDE_VALUE"
#     set:
#       Name2: "REPLACE_WITH_NAME2_VALUE"
#     overrideCell: no
```

- `Kundennummer`: required, quoted, non-empty text. Quoting preserves leading zeros.
- `Kunde`: required, non-empty text that exactly matches the workbook value.
- `set`: required mapping of one or both supported columns, `Name1` and `Name2`, to non-empty text values.
- `overrideCell`: optional YAML boolean (`yes` or `no`); defaults to `no`.

Add multiple corrections as separate, equally indented items under `rules`. Do not add other keys to a rule.

### REVIEW: Pre-Generation Customer Name Review in Excel

After enrichment and before generating contracts, review whether the contract name (`Kunde`) looks consistent with the SAP names (`Name1` and `Name2`). Add a helper column to the `uniform` sheet. The formula below returns a score from 0 to 1: the share of distinctive words in `Kunde` also found in `Name1` or `Name2`.

This formula requires Excel 365 functions (`LET`, `TEXTSPLIT`, `UNIQUE`, `FILTER`, and `XMATCH`). With the current column layout, enter it on row 2 and fill down:

```excel
=IF(OR([@Kunde]="";AND([@Name1]="";[@Name2]=""));" ";IFERROR(LET(k;LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B2;"-";" ");".";" ");"/";" ");",";" "));n;" "&LOWER(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE([@Name1]&" "&[@Name2];"-";" ");".";" ");"/";" ");",";" "))&" ";t;UNIQUE(TEXTSPLIT(k;" "));useful;FILTER(t;(LEN(t)>=4)*ISNA(XMATCH(t;{"klinik";"klinikum";"krankenhaus";"gmbh";"ggmbh";"stadt";"und";"der";"des";"von";"am";"im"})));SUM(--ISNUMBER(SEARCH(" "&useful&" ";n)))/COUNTA(useful));0))
```

Use the score only to prioritize manual review before running `src/generate_contracts_separate.py`. Investigate low or zero scores and resolve any incorrect customer details in the source data before proceeding. A low or zero score does not necessarily mean the mapping is wrong: names may use abbreviations, aliases, or different spellings. Do not use this score to replace the `Kundennummer` join key. If Excel uses semicolons as formula separators, replace the commas accordingly.

Every completed review must be documented: record what was checked, what was corrected, who performed the review, and when it was completed. After the review is finished, the reviewed cell or worksheet must be protected so that the validated result cannot be changed accidentally. **THE CELL NEEDS TO BE PROTECTED**.

---

## Automated Contract Generation

### Script

`src/generate_contracts_separate.py`

The enriched workbook serves as the source for document generation.

The automation performs the following steps:

- Groups all records by the `Kunde` customer name; matching rows do not need to be adjacent
- Reads the associated customer address information
- Preserves all product and application entries
- Populates the Word template automatically
- Creates one EVB-IT contract per customer

### Customer grouping and customer numbers

The current contract-generation script groups rows by the `Kunde` value, not by `Kundennummer`. It collects matching rows from anywhere in the workbook, so they do not need to be next to each other.

Within each `Kunde` group, the script collects the distinct, non-empty `Kundennummer` values. If there is exactly one, it uses that number in the contract. If there are multiple numbers, it writes them with a `PRÜFEN:` prefix for review. Consequently, the same customer name with multiple customer numbers is generated as one contract flagged for review, while the same customer number under different `Kunde` names is generated as separate contracts.

Rows whose `Kundennummer` is exactly `-` are intentional skips: SAP lookup, joining, and contract generation ignore them without reporting a missing-customer error. A `?` anywhere in a customer number is a team-confirmation marker. SAP lookup and the generated contract use the number with all `?` characters removed; the SAP export retains the marked value so the join still matches. The contract-generation log flags the customer as requiring team confirmation.

If contracts should instead be grouped by customer number, the grouping key in the script must be changed from `Kunde` to `Kundennummer`.

Example:

```text
<example customer number>
├── Monitor              18
├── DBMS               5
└── DeviceManagerplus     2
```

This structure is maintained within the generated contract document.

Generated contracts are saved under `output/{Innendienst}_({ID-Vert.})`, with
spaces removed from both folder-name values (for example,
`Max M` and `Alice M` become `MaxM_(AliceM)`). If rows for one
customer contain different Innendienst/representative assignments, the
generator logs a warning and saves that contract directly in `output` for
review.

---

## Automated Document Formatting

### Script

`src/beautify_contracts.py`

After contract generation, the created Word documents undergo a second formatting step.
The formatter scans `output` and its subfolders for Word documents.

This module applies standardized formatting rules such as:

- Arial font normalization
- Consistent table styling
- Standardized 8pt table font size
- Corporate layout adjustments
- Product name formatting
- Superscript "+" formatting for branded applications

Examples:

- Monitor⁺
- DeviceManager⁺

Separating formatting from content generation improves maintainability and simplifies future enhancements.

---

## Recommended Execution Sequence

### 1. Retrieve Missing Customer Data from SAP

```powershell
python src\sap_client.py
```

Output:

```text
res/sap_customer_export.csv
```

### 2. Enrich Contract Data

```powershell
python src\join_python.py
```

Output:

```text
res/EVB_IT_Vertragsuebersicht_Enriched.xlsx
logs/join_python_<timestamp>.log
```

Each active Python script creates a timestamped `.log` file in `logs/`. Each entry includes a timestamp, and unhandled errors are recorded with their traceback. Contract-generation summary reports are saved in the same folder. `src/sap_client.py` preserves existing CSV rows and retrieves only customer numbers not already listed there.

### 3. Compare the Source and Enriched Workbooks

```powershell
python src\compare_excel.py
```

By default, the script compares the configured `comparison_first` and
`comparison_second` workbooks in `config.yaml`, using the `uniform` sheet.
You can optionally pass two workbook paths and a sheet name:

```powershell
python src\compare_excel.py "res\first.xlsx" "res\second.xlsx" --sheet uniform
```

Rows are matched by content across the entire sheet, regardless of their
position, with columns matched by header. `Anmerkungen` is excluded from the
content comparison. Each occurrence can be matched only once; equal rows are
counted as the smaller number of occurrences in the two files.

The report identifies `Erste Datei` and `Zweite Datei` by their paths. For each
content group with unequal occurrence counts, it shows the counts, physical
row numbers, and available customer identifiers. Unmatched occurrences are warnings
if any row in the group has `OK` in `Anmerkungen` (case-insensitive), or if
the content occurs more than once in either file. Other unmatched occurrences
are regular deviations. The deviation and warning totals count unmatched
occurrences, so changing one row can count as one missing and one added row.
Warnings are included in `Abweichende Zeilen` and counted separately.

For a possible column-by-column comparison, unmatched content groups must have
the same nonempty `Kundennummer` and `Kunde`. Groups are paired only when they
are each other's unique closest match, measured by the number of differing
columns; this is repeated for the remaining groups. Only changed columns are
shown as `old -> new`. These are possible matches, not proof that a row was
edited. Ambiguous or unpaired groups are reported without a column comparison.
This diagnostic pairing does not change the occurrence counts or warnings.

Each run writes a timestamped report to `logs/compare_excel_<timestamp>.log`.
The report includes the row differences, totals, and warnings.

### 4. Review Customer Name Mappings

Open `res/EVB_IT_Vertragsuebersicht_Enriched.xlsx` and use the Excel formula in the [pre-generation review section](#pre-generation-customer-name-review-in-excel) to flag and manually check possible name mismatches. Resolve any incorrect customer details before continuing.

### 5. Generate Contracts

```powershell
python src\generate_contracts_separate.py
```

### 6. Apply Final Formatting

```powershell
python src\beautify_contracts.py
```

---

## Python Environment Setup

### Install Required Libraries

```powershell
python -m pip install pandas openpyxl python-docx docxtpl pywin32 PyYAML
```

### Upgrade Existing Libraries

```powershell
pip install --upgrade pandas openpyxl python-docx docxtpl pywin32 PyYAML
```

### Verify Installed Versions

```powershell
pip list
```

or

```powershell
pip show pandas openpyxl python-docx docxtpl pywin32 PyYAML
```

---

## Virtual Environment

### Windows PowerShell

```powershell
.\.venv\Scripts\Activate.ps1
```

### Windows Command Prompt

```cmd
.venv\Scripts\activate.bat
```

---

## Pre-Execution Checklist

Before running the automation:

- ✅ Virtual environment activated
- ✅ SAP GUI running
- ✅ SAP GUI Scripting enabled
- ✅ Required Python libraries installed
- ✅ Libraries updated
- ✅ Contract overview workbook available
- ✅ Word template available
- ✅ SAP customer extraction completed
- ✅ Contract enrichment completed successfully
- ✅ Product mappings verified
- ✅ Output directory accessible

Only after all validation steps have been completed should the contract creation process be executed.

---

## Business Benefits

The automation delivers several operational advantages:

- Reduced manual contract preparation effort
- Improved customer master data accuracy
- Elimination of repetitive copy-and-paste activities
- Standardized EVB-IT contract generation
- Faster processing of large customer migrations
- Improved scalability and maintainability
- Reusable framework for future AIS automation initiatives

---

## Technologies Used

- Python
- Pandas
- OpenPyXL
- Python-Docx
- DocxTpl
- SAP GUI Scripting
- PyWin32
- Microsoft Excel
- Microsoft Word

---

## Acknowledgements

Special thanks to all colleagues and stakeholders who supported the EVB-IT migration project through testing, validation, process expertise, and business feedback during the development of this automation solution.

---

**Agha M. Aslam**  
**Intern | AIS-Innendiest Automation & Process Digitalization**  
**B. Braun 2026**