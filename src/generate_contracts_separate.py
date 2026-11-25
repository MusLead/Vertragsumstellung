"""
THIS VERSIONS ALLOWED EACH GROUPING OF PRODUCTS TO BE GENERATED AS A SEPARATE ROWS.
Create contracts from Excel; product descriptions and row colors come from Catalog.
"""

import os
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

from logging_config import setup_logger
from project_config import get_configured_path

logger, LOG_FILE = setup_logger(__file__)

import pandas as pd
from docxtpl import DocxTemplate, RichText


DEBUG = False  # Set to True to skip contract generation and only log the process
EXCEL = get_configured_path("enriched_workbook")
TEMPLATE = get_configured_path("contract_template")
OUTPUT = get_configured_path("contracts_output")
os.makedirs(OUTPUT, exist_ok=True)

CUSTOMER_FIRST_NAME_COLUMN = "Name1"
CUSTOMER_LAST_NAME_COLUMN = "Name2"
CUSTOMER_STREET_COLUMN = "Street"
CUSTOMER_HOUSE_NUMBER_COLUMN = "HouseNumber"
CUSTOMER_POSTAL_CODE_COLUMN = "PostalCode"
CUSTOMER_CITY_COLUMN = "City"
CUSTOMER_PROD_TYPE_COLUMN = "Typ"
CUSTOMER_QUANTITY_COLUMN = "Menge"
# CUSTOMER_SPECIAL_PRICE_COLUMN = "SpecialPrice"
CUSTOMER_PRICE_COLUMN = "Price"
CUSTOMER_DISCOUNT_COLUMN = "Rabatt"
CUSTOMER_SUBTOTAL_COLUMN = "Subtotal"
INNENDIENST_COLUMN = "Innendienst"
INNENDIENST_REPRESENTATIVE_COLUMN = "ID-Vert."

# Excel product name -> Catalog article number
PRODUCTS = {
    'OneView': '876205-4A',
    'ULM.Web': '876204-4A'
}

REQUIRED_UNIFORM_COLUMNS = (
    'Kundennummer',
    'Kunde',
    CUSTOMER_FIRST_NAME_COLUMN,
    CUSTOMER_STREET_COLUMN,
    CUSTOMER_HOUSE_NUMBER_COLUMN,
    CUSTOMER_POSTAL_CODE_COLUMN,
    CUSTOMER_CITY_COLUMN,
    CUSTOMER_PROD_TYPE_COLUMN,
    CUSTOMER_QUANTITY_COLUMN,
    # CUSTOMER_SPECIAL_PRICE_COLUMN,
    CUSTOMER_PRICE_COLUMN,
    CUSTOMER_DISCOUNT_COLUMN,
    INNENDIENST_COLUMN,
    INNENDIENST_REPRESENTATIVE_COLUMN,
)


def safe_folder_component(value, remove_whitespace=False):
    text = str(value).strip()
    if remove_whitespace:
        text = "".join(text.split())

    text = "".join(
        "_"
        if ord(character) < 32 or character in '<>:"/\\|?*'
        else character
        for character in text
    )
    return text.strip(" .")


def validate_workbook_columns(table_columns):
    missing_columns = {
        'uniform': [
            column
            for column in REQUIRED_UNIFORM_COLUMNS
            if column not in table_columns['uniform']
        ],
    }

    if len(table_columns['Catalog']) < 5:
        missing_columns['Catalog'] = [
            'at least 5 columns (article number through price)'
        ]

    missing_columns = {
        table: columns
        for table, columns in missing_columns.items()
        if columns
    }

    if not missing_columns:
        return

    details = '\n'.join(
        f"  {table}: {columns!r}"
        for table, columns in table_columns.items()
    )
    missing = '\n'.join(
        f"  {table}: {', '.join(columns)}"
        for table, columns in missing_columns.items()
    )
    message = (
        "Workbook column validation failed.\n"
        f"Missing required columns:\n{missing}\n"
        f"Columns found in the workbook:\n{details}"
    )
    logger.error("%s", message)
    raise SystemExit(message)


def stuck(value):
    return f"{int(value)} Stk."

def toPercent(value):
    try:
        number = float(value)
        if 0 < number < 1: # e.g. 0.15 -> 15%
            number *= 100
        return f"{int(number)}%"
    except (ValueError, TypeError):
        return "0%"


CENT = Decimal("0.01")
ONE = Decimal("1")
HUNDRED = Decimal("100")


def to_decimal(value):
    text = (
        str(value)
        .strip()
        .replace("€", "")
        .replace("\xa0", "")
        .replace(" ", "")
    )
    if "," in text: # example 1.234,56 -> 1234.56
        text = text.replace(".", "").replace(",", ".")
    return Decimal(text)


def euros(value):
    return (
        f"{value:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
        + " €"
    )

created_contracts = []
skipped_contracts = []


# --------------------------------------------------
# READ EXCEL
# --------------------------------------------------

try:
    logger.info("Loading contract and catalog data from %s", EXCEL)
    excel_book = pd.ExcelFile(EXCEL, engine='openpyxl')
except PermissionError:
    logger.exception("Could not open Excel workbook %s", EXCEL)
    raise SystemExit(
        "The Microsoft file could not be opened. It might still be open. "
        "Please close this Excel file first. Thank you! :)"
    ) from None

with excel_book as book:

    intentional_skip_count = 0
    data = pd.read_excel(
        book,
        'uniform',
        dtype={
            'Kundennummer': str,
            CUSTOMER_POSTAL_CODE_COLUMN: str
        }
    ).astype(object).fillna('')

    catalog = pd.read_excel(
        book,
        'Catalog',
        header=None
    ).astype(object).fillna('')

    table_columns = {
        'uniform': list(data.columns),
        'Catalog': (
            list(catalog.iloc[0])
            if not catalog.empty
            else list(catalog.columns)
        ),
    }
    validate_workbook_columns(table_columns)

    intentional_skip_rows = data["Kundennummer"].astype(str).str.strip().eq("-")
    if intentional_skip_rows.any():
        intentional_skip_count = int(intentional_skip_rows.sum())
        logger.info(
            "Ignoring %s row(s) with Kundennummer '-' as intentional skips",
            intentional_skip_count,
        )
        data = data.loc[~intentional_skip_rows].copy()

    sheet = book.book['Catalog']

    #--------------------------------------------------
    # Read catalog data into articles dictionary
    #--------------------------------------------------
    
    articles = {}

    for i, row in catalog.iterrows():

        nr = str(row.iloc[0]).strip()

        if nr not in PRODUCTS.values():
            continue

        # Each article block continues until next article number
        end = i + 1

        while (
            end < len(catalog)
            and not str(catalog.iloc[end, 0]).strip()
        ):
            end += 1

        desc = "\n".join(
            str(x).strip()
            for x in catalog.iloc[i:end, 3]
            if str(x).strip()
        )

        fill = sheet.cell(i + 1, 4).fill
        color = fill.fgColor

        bg = (
            color.rgb[-6:]
            if (
                fill.patternType == "solid"
                and color.type == "rgb"
                and isinstance(color.rgb, str)
            )
            else "FFFFFF"
        )

        articles[nr] = {
            "nr": nr,
            "name": str(row.iloc[1]).strip(),
	        "apps": str(row.iloc[2]).strip(),
            "desc": desc,
            "price": to_decimal(row.iloc[4]),
            "bg": bg
        }


# --------------------------------------------------
# CREATE CONTRACTS
# --------------------------------------------------

for customer, group in data.groupby('Kunde', sort=False):

    customer_data = group.iloc[0]
    raw_numbers = (
        group["Kundennummer"]
        .astype(str)
        .str.strip()
    )
    needs_team_confirmation = raw_numbers.str.contains("?", regex=False).any()
    cleaned_numbers = raw_numbers.str.replace("?", "", regex=False).str.strip()

    if needs_team_confirmation:
        logger.warning(
            "TEAM CONFIRMATION REQUIRED for %s (Kundennummer: %s)",
            customer,
            " / ".join(sorted(set(cleaned_numbers) - {""})),
        )

    #--------------------------------------------------
    # Check if all products in the group have status OK
    # Check if required customer fields are present
    #--------------------------------------------------

    statuses = (
        group['Status']
        .astype(str)
        .str.strip()
        .str.upper()
        .unique()
        .tolist()
    )

    if group['Status'] \
            .astype(str) \
            .str.strip() \
            .str.upper() \
            .eq('ERROR') \
            .all():

        reason = f"Status not OK ({', '.join(statuses)})"

        skipped_contracts.append({
            "customer": customer,
            "reason": reason,
            "needs_team_confirmation": needs_team_confirmation,
        })

        print(f"SKIPPED: {customer} -> {reason}")
        logger.warning("Skipped %s: %s", customer, reason)
        continue # CONTRACT WILL NOT BE GENERATED IF STATUS IS NOT OK

    required_customer_fields = (
        CUSTOMER_FIRST_NAME_COLUMN,
        CUSTOMER_STREET_COLUMN,
        # CUSTOMER_HOUSE_NUMBER_COLUMN,
        CUSTOMER_POSTAL_CODE_COLUMN,
        CUSTOMER_CITY_COLUMN,
    )

    missing_customer_fields = [
        field
        for field in required_customer_fields
        if (
            pd.isna(customer_data[field])
            or not str(customer_data[field]).strip()
        )
    ]

    has_sap_log_error = (
        group['SAP-LOG']
        .astype(str)
        .str.strip()
        .ne('')
        .any()
    )

    if missing_customer_fields or has_sap_log_error:
        reason = (
            "ERROR retrieving data from SAP"
            if has_sap_log_error
            else (
                "Missing required customer data: "
                f"{', '.join(missing_customer_fields)}"
            )
        )

        skipped_contracts.append({
            "customer": customer,
            "reason": reason,
            "needs_team_confirmation": needs_team_confirmation,
        })

        print(f"SKIPPED: {customer} -> {reason}")
        logger.warning("Skipped %s: %s", customer, reason)
        continue # CONTRACT WILL NOT BE GENERATED IF REQUIRED CUSTOMER DATA IS MISSING OR SAP LOG ERROR EXISTS

    assignments = {
        (
            str(innendienst).strip(),
            "".join(str(representative).split()),
        )
        for innendienst, representative in zip(
            group[INNENDIENST_COLUMN],
            group[INNENDIENST_REPRESENTATIVE_COLUMN],
        )
    }

    output_directory = OUTPUT
    if len(assignments) > 1:
        logger.warning(
            "Inconsistent Innendienst assignment for %s (%s); "
            "saving contract in output root",
            customer,
            sorted(assignments),
        )
    else:
        innendienst, representative = next(iter(assignments))
        innendienst_folder = safe_folder_component(
            innendienst,
            remove_whitespace=True,
        )
        representative_folder = safe_folder_component(
            representative,
            remove_whitespace=True,
        )
        if innendienst_folder and representative_folder:
            output_directory = os.path.join(
                OUTPUT,
                f"{innendienst_folder}_({representative_folder})",
            )
        else:
            logger.warning(
                "Missing Innendienst or ID-Vert. for %s; "
                "saving contract in output root",
                customer,
            )
    
    #--------------------------------------------------
    # Determine customer number to use in contract
    #--------------------------------------------------

    numbers = sorted(set(cleaned_numbers) - {""})

    if needs_team_confirmation and not numbers:
        reason = "No usable Kundennummer remains after removing '?' markers"
        skipped_contracts.append({
            "customer": customer,
            "reason": reason,
            "needs_team_confirmation": True,
        })
        logger.warning("Skipped %s: %s", customer, reason)
        continue

    customer_number = (
        numbers[0]
        if len(numbers) == 1 # Number is unique, use it
        # If there are multiple numbers, for one customer, REVISON NEED: This is a problem, but we will still generate the contract with all numbers listed for review
        else "PRÜFEN: " + " / ".join(numbers) 
    )

    #--------------------------------------------------
    # For each customer, create table rows for each product in the group
    #--------------------------------------------------
    rows = []
    total = Decimal("0.00")

    for _, row_data in group.iterrows():

        product = str(row_data[CUSTOMER_PROD_TYPE_COLUMN]).strip() 

        if (
            product not in PRODUCTS
            or PRODUCTS[product] not in articles
        ):
            raise ValueError(
                f"Unknown product mapping: {product}"
            )

        article = articles[PRODUCTS[product]]
        
        amount = int(float(row_data[CUSTOMER_QUANTITY_COLUMN]))
        
        # special_price = str(
        #     row_data[CUSTOMER_SPECIAL_PRICE_COLUMN]
        # ).strip()
        # price = (
        #     to_decimal(special_price)
        #     if special_price
        #     else article["price"]
        # )

        # discount_text = str(
        #     row_data[CUSTOMER_DISCOUNT_COLUMN]
        # ).strip()
        # discount = to_decimal(discount_text) if discount_text else Decimal("0")
        
        # price = (
        #     price * (ONE - discount / HUNDRED) # if discount is 0%, then price * (1 - 0/100) = price * 1 = price
        # ).quantize(CENT, rounding=ROUND_HALF_UP)
        
        # subtotal = (price * amount).quantize(
        #     CENT,
        #     rounding=ROUND_HALF_UP,
        # )
        
        # total += subtotal

        # price_text = (
        #     f"{euros(Decimal(special_price) if special_price.split() else article["price"])}"
        #     + (
        #         f"\n(abzgl. {int(discount)}%)"
        #         if discount > 0
        #         else ""
        #     )
        # )

        #--------------------------------------------------
        # Retrieve price, discount, and subtotal from the Excel data for each product
        #--------------------------------------------------
        discount_text = str(row_data[CUSTOMER_DISCOUNT_COLUMN]).strip()
        price_text = euros(to_decimal(row_data[CUSTOMER_PRICE_COLUMN])) + (
            f"\n(abzgl. {toPercent(discount_text)})"
            if discount_text and discount_text != "0"
            else ""
        )
        subtotal_text = str(row_data[CUSTOMER_SUBTOTAL_COLUMN]).strip()

        subtotal = to_decimal(subtotal_text)
        total += subtotal
        #--------------------------------------------------

        rows.append({
            "nr": article["nr"],
            "name": article["name"],
            "apps": article["apps"],
            "desc": article["desc"],
            "bg": article["bg"],
            "price": price_text,
            "quantity": stuck(amount),
            "total": euros(subtotal)
        })

    #--------------------------------------------------
    # Render the document
    #--------------------------------------------------

    customer_name = (
        f"{customer_data[CUSTOMER_FIRST_NAME_COLUMN].strip()}\n"
        f"{customer_data[CUSTOMER_LAST_NAME_COLUMN].strip()}"
        if customer_data[CUSTOMER_LAST_NAME_COLUMN].strip()
        else f"{customer_data[CUSTOMER_FIRST_NAME_COLUMN].strip()}"
    )
    postal_code = str(customer_data[CUSTOMER_POSTAL_CODE_COLUMN]).strip()
    if postal_code.isdigit():
        postal_code = postal_code.zfill(5)

    if DEBUG : continue

    doc = DocxTemplate(TEMPLATE)
    doc.render({
        # "VERTRAGSPARTNER": customer,
        "KUNDENNUMMER": customer_number,
        "NAME": customer_name,
        "STRASSE": (
            f"{customer_data[CUSTOMER_STREET_COLUMN]} "
            f"{customer_data[CUSTOMER_HOUSE_NUMBER_COLUMN]}"
        ),
        "PLZ_ORT": (
            f"{postal_code} {customer_data[CUSTOMER_CITY_COLUMN]}"
        ),
        "LIZENZGEBUEHR": euros(total),
        "produkte": rows # FOR THE TABLE IN THE CONTRACT
    })

    filename = (
        f"{safe_folder_component(customer_number)}_"
        f"{safe_folder_component(customer)}"
        "__REVIEW.docx"
    )

    output_file = os.path.join(output_directory, filename)
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    relative_output_file = os.path.relpath(output_file, OUTPUT)

    doc.save(output_file)

    created_contracts.append({
        "customer": customer,
        "kundennummer": customer_number,
        "products": len(rows),
        "file": relative_output_file,
        "needs_team_confirmation": needs_team_confirmation,
    })

    logger.info(
        "Created contract for %s: %s",
        customer,
        output_file,
    )
    print(
        f"CREATED: {customer} "
        f"({len(rows)} products, "
        f"Kundennummer: {customer_number})"
    )


# --------------------------------------------------
# LOG FILE
# --------------------------------------------------

summary_file = LOG_FILE.with_name(f"{LOG_FILE.stem}_summary.txt")

with open(
    summary_file,
    "w",
    encoding="utf-8"
) as log:

    log.write(
        "CONTRACT GENERATION LOG\n"
    )

    log.write(
        "=" * 80 + "\n\n"
    )

    log.write(
        f"Generated on: "
        f"{datetime.now():%d.%m.%Y %H:%M:%S}\n\n"
    )

    log.write(
        f"Contracts created: "
        f"{len(created_contracts)}\n"
    )

    log.write(
        f"Customers skipped: "
        f"{len(skipped_contracts)}\n\n"
    )
    log.write(
        f"Intentional '-' row skips: "
        f"{intentional_skip_count}\n\n"
    )

    log.write(
        "CREATED CONTRACTS\n"
    )

    log.write(
        "-" * 80 + "\n"
    )

    for item in created_contracts:

        confirmation_note = (
            " | TEAM CONFIRMATION REQUIRED"
            if item["needs_team_confirmation"]
            else ""
        )
        log.write(
            f"{item['customer']} | "
            f"Kundennummer: "
            f"{item['kundennummer']} | "
            f"Products: "
            f"{item['products']} | "
            f"File: "
            f"{item['file']}"
            f"{confirmation_note}\n"
        )

    log.write("\n")

    log.write(
        "SKIPPED CUSTOMERS\n"
    )

    log.write(
        "-" * 80 + "\n"
    )

    if skipped_contracts:

        for item in skipped_contracts:

            confirmation_note = (
                " | TEAM CONFIRMATION REQUIRED"
                if item["needs_team_confirmation"]
                else ""
            )
            log.write(
                f"{item['customer']} | "
                f"Reason: "
                f"{item['reason']}"
                f"{confirmation_note}\n"
            )

    else:

        log.write("None\n")


print()
print("=" * 80)
print("SUMMARY")
print("=" * 80)
print(
    f"Contracts created: "
    f"{len(created_contracts)}"
)
print(
    f"Customers skipped: "
    f"{len(skipped_contracts)}"
)
print("Skipped customers:")
if skipped_contracts:
    for item in skipped_contracts:
        print(
            f"- {item['customer']} | "
            f"Reason: {item['reason']}"
        )
else:
    print("None")
print(
    f"Log file: "
    f"{LOG_FILE}"
)
print(f"Summary file: {summary_file}")
logger.info(
    "Run completed: %s contracts created, %s customers skipped",
    len(created_contracts),
    len(skipped_contracts),
)