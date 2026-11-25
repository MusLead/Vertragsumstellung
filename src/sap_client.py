import csv
import os
import time
import pandas as pd
from logging_config import setup_logger
from project_config import get_configured_path


logger, LOG_FILE = setup_logger(__file__)


EXCEL_FILE = get_configured_path("contract_workbook")
OUTPUT_CSV = get_configured_path("sap_export")
CSV_COLUMNS = (
    "Kundennummer",
    "Name1",
    "Name2",
    "Street",
    "HouseNumber",
    "PostalCode",
    "City",
    "SAP-LOG",
)

BASE_PATH = (
    "wnd[0]/usr/"
    "subSUBTAB:SAPLATAB:0100/"
    "tabsTABSTRIP100/tabpTAB01/"
    "ssubSUBSC:SAPLATAB:0201/"
    "subAREA1:SAPMF02D:7111/"
    "subADDRESS:SAPLSZA1:0300/"
    "subCOUNTRY_SCREEN:SAPLSZA1:0301/"
)


def normalize_customer_number(value):
    return str(value).strip()


def customer_number_for_sap(value):
    return normalize_customer_number(value).replace("?", "").strip()


def load_customer_numbers(excel_file):
    logger.info("Loading customer numbers from %s", excel_file)
    try:
        df = pd.read_excel(
            excel_file,
            sheet_name="uniform",
            dtype={"Kundennummer": str},
        )
    except PermissionError:
        logger.exception("Could not open Excel workbook %s", excel_file)
        raise SystemExit(
            "Could not open the Excel workbook. Please close the Excel table "
            "if it is open, then run this script again."
        ) from None

    customer_numbers = (
        df["Kundennummer"]
        .dropna()
        .map(normalize_customer_number)
        .loc[lambda numbers: numbers.ne("")]
    )
    skipped_rows = customer_numbers.eq("-").sum()
    if skipped_rows:
        logger.info(
            "Ignoring %s row(s) with Kundennummer '-' as intentional skips",
            skipped_rows,
        )

    return customer_numbers.loc[lambda numbers: numbers.ne("-")].unique().tolist()


def load_existing_customer_numbers(csv_file):
    if not os.path.exists(csv_file) or os.path.getsize(csv_file) == 0:
        return set()

    with open(csv_file, encoding="utf-8-sig", newline="") as file:
        reader = csv.DictReader(file, delimiter=";")
        if not reader.fieldnames or "Kundennummer" not in reader.fieldnames:
            raise ValueError(
                f"{csv_file} must contain a 'Kundennummer' column"
            )
        return {
            normalize_customer_number(row["Kundennummer"])
            for row in reader
            if row.get("Kundennummer")
            and normalize_customer_number(row["Kundennummer"])
        }


def build_customer_queue(customer_numbers, existing_customer_numbers):
    queue = []
    for customer in customer_numbers:
        normalized_customer = normalize_customer_number(customer)
        if (
            normalized_customer
            and normalized_customer != "-"
            and normalized_customer not in existing_customer_numbers
        ):
            queue.append(normalized_customer)
    return queue


def append_result_to_csv(result, csv_file=OUTPUT_CSV):
    file_exists = os.path.exists(csv_file)
    needs_header = not file_exists or os.path.getsize(csv_file) == 0
    encoding = "utf-8-sig" if needs_header else "utf-8"

    with open(csv_file, "a", encoding=encoding, newline="") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=CSV_COLUMNS,
            delimiter=";",
        )
        if needs_header:
            writer.writeheader()
        writer.writerow({column: result.get(column, "") for column in CSV_COLUMNS})


def stop_for_sap_setup(message, reason):
    logger.error("SAP is not ready: %s", reason)
    raise SystemExit(message)


def connect_to_sap():
    import pywintypes
    import win32com.client

    logger.info("Connecting to SAP GUI")
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine
    except pywintypes.com_error:
        stop_for_sap_setup(
            "Could not connect to SAP GUI. Please turn on SAP Client, Choose a DEA System, log in and then try "
            "again.",
            "SAP GUI scripting object is unavailable",
        )

    try:
        connection = application.Children(0)
    except pywintypes.com_error:
        stop_for_sap_setup(
            "SAP Client is open, but SAP DEA system is not connected. Please connect and log in with your username and password, then try again.",
            "no SAP connections are open",
        )

    try:
        session = connection.Children(0)
    except pywintypes.com_error:
        stop_for_sap_setup(
            "No SAP session is open. Please connect to SAP and log in with your "
            "username and password, then try again.",
            "the active SAP connection has no sessions",
        )

    try:
        session.findById("wnd[0]/usr/txtRSYST-BNAME")
        session.findById("wnd[0]/usr/pwdRSYST-BCODE")
    except pywintypes.com_error:
        pass
    else:
        stop_for_sap_setup(
            "Please log in to the SAP application with your username and "
            "password, then try again.",
            "SAP is displaying the logon screen",
        )

    logger.info("Connected to SAP GUI")
    return session


def retrieve_customer(session, customer):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/nVD03"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)

    session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = customer
    session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "00"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(2)

    result = {
        "Kundennummer": customer,
        "Name1": session.findById(
            BASE_PATH + "txtADDR1_DATA-NAME1"
        ).text,
        "Name2": session.findById(
            BASE_PATH + "txtADDR1_DATA-NAME2"
        ).text,
        "Street": session.findById(
            BASE_PATH + "txtADDR1_DATA-STREET"
        ).text,
        "HouseNumber": session.findById(
            BASE_PATH + "txtADDR1_DATA-HOUSE_NUM1"
        ).text,
        "PostalCode": session.findById(
            BASE_PATH + "txtADDR1_DATA-POST_CODE1"
        ).text,
        "City": session.findById(
            BASE_PATH + "txtADDR1_DATA-CITY1"
        ).text,
        "SAP-LOG": "",
    }

    session.findById("wnd[0]").sendVKey(3)
    time.sleep(1)
    return result


def main():
    customer_numbers = load_customer_numbers(EXCEL_FILE)
    existing_customer_numbers = load_existing_customer_numbers(OUTPUT_CSV)
    customer_queue = build_customer_queue(
        customer_numbers,
        existing_customer_numbers,
    )
    lookup_queue = []
    skipped_without_lookup = 0
    for customer in customer_queue:
        lookup_customer = customer_number_for_sap(customer)
        if not lookup_customer or lookup_customer == "-":
            message = (
                "No usable Kundennummer remains after removing '?' "
                "confirmation markers"
            )
            logger.error(
                "Cannot retrieve customer %s from SAP: %s",
                customer,
                message,
            )
            append_result_to_csv({
                "Kundennummer": customer,
                "Name1": "",
                "Name2": "",
                "Street": "",
                "HouseNumber": "",
                "PostalCode": "",
                "City": "",
                "SAP-LOG": message,
            })
            skipped_without_lookup += 1
            continue

        if "?" in customer:
            logger.warning(
                "Customer %s needs team confirmation; looking up %s in SAP",
                customer,
                lookup_customer,
            )
        lookup_queue.append((customer, lookup_customer))

    print(f"Found {len(customer_numbers)} unique customer numbers")
    print(f"Already in CSV: {len(existing_customer_numbers)}")
    print(f"Queued for SAP retrieval: {len(lookup_queue)}")
    logger.info(
        "Found %s workbook customers; %s already in CSV; %s queued for SAP",
        len(customer_numbers),
        len(existing_customer_numbers),
        len(lookup_queue),
    )

    if not lookup_queue:
        print(f"No new customer data to retrieve. Existing export: {OUTPUT_CSV}")
        if skipped_without_lookup:
            logger.warning(
                "%s customer(s) could not be looked up because no usable "
                "number remained after removing '?' markers",
                skipped_without_lookup,
            )
        else:
            logger.info("No customers require SAP retrieval")
        return

    session = connect_to_sap()

    for customer, lookup_customer in lookup_queue:
        try:
            print(f"Processing: {customer}")
            logger.info("Processing customer %s", customer)
            result = retrieve_customer(session, lookup_customer)
            result["Kundennummer"] = customer
            print(f"SUCCESS: {customer} -> {result['Name1']}")
            logger.info("Retrieved customer %s: %s", customer, result["Name1"])
        except Exception as error:
            print(f"ERROR: {customer}")
            print(error)
            logger.exception("Failed to retrieve customer %s", customer)
            result = {
                "Kundennummer": customer,
                "Name1": "",
                "Name2": "",
                "Street": "",
                "HouseNumber": "",
                "PostalCode": "",
                "City": "",
                "SAP-LOG": str(error),
            }

        append_result_to_csv(result)

    appended_records = len(lookup_queue) + skipped_without_lookup
    logger.info("Appended %s records to %s", appended_records, OUTPUT_CSV)
    print()
    print("=" * 60)
    print(f"Export updated: {OUTPUT_CSV}")
    print(f"Records processed: {appended_records}")
    print(f"Log file: {LOG_FILE}")
    logger.info("Run completed successfully")


if __name__ == "__main__":
    main()
