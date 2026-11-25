import os

from logging_config import setup_logger
from project_config import get_configured_path

logger, LOG_FILE = setup_logger(__file__)

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.exceptions import PackageNotFoundError
from docx.shared import Pt, RGBColor

OUTPUT_FOLDER = get_configured_path("contracts_output")

def set_run_font(run, size=None):

    run.font.name = "Arial"

    # Needed for Word compatibility
    try:
        run._element.rPr.rFonts.set(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii',
            'Arial'
        )
        run._element.rPr.rFonts.set(
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi',
            'Arial'
        )
    except:
        pass

    if size is not None:
        run.font.size = Pt(size)

def set_cell_background(cell, fill_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)

    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_color)


def set_text_color_for_runs(cell, text_color):
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor.from_string(text_color)


def format_table(table):

    table.autofit = True

    for row_index, row in enumerate(table.rows):

        row.height = None
        row.height_rule = None

        if row_index == 0:
            fill_color = "4ba9a4"
            text_color = "FFFFFF"
        else:
            fill_color = "FFFFFF" if row_index % 2 == 1 else "daeeed"
            text_color = "000000"

        for cell in row.cells:
            set_cell_background(cell, fill_color)

            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    set_run_font(run, size=8)
                    run.font.color.rgb = RGBColor.from_string(text_color)

            for nested_table in cell.tables:
                format_table(nested_table)

def split_plus_runs(paragraph):

    text = paragraph.text

    if "plus" not in text.lower():
        return

    font_size = next(
        (
            run.font.size
            for run in paragraph.runs
            if run.font.size is not None and not run.font.superscript
        ),
        None
    )

    text_lower = text.lower()

    idx = text_lower.find("plus")

    if idx == -1:
        return

    before = text[:idx]
    plus_text = text[idx:idx+4]
    after = text[idx+4:]

    paragraph.clear()

    run1 = paragraph.add_run(before)
    run1.bold = True
    run1.font.size = font_size

    run2 = paragraph.add_run(plus_text)
    run2.bold = True
    run2.font.size = Pt(7)
    run2.font.superscript = True

    hybrid_fleet_idx = after.lower().find("hybrid fleet")
    if hybrid_fleet_idx == -1:
        run3 = paragraph.add_run(after)
        run3.font.size = font_size
        return

    before_hybrid_fleet = after[:hybrid_fleet_idx]
    hybrid_fleet = after[
        hybrid_fleet_idx:hybrid_fleet_idx + len("Hybrid Fleet")
    ]
    after_hybrid_fleet = after[
        hybrid_fleet_idx + len("Hybrid Fleet"):
    ]

    run3 = paragraph.add_run(before_hybrid_fleet)
    run3.font.size = font_size

    run4 = paragraph.add_run(hybrid_fleet)
    run4.bold = True
    run4.font.size = font_size

    run5 = paragraph.add_run(after_hybrid_fleet)
    run5.font.size = font_size

def process_table(table):

    for row in table.rows:

        for cell in row.cells:

            for paragraph in cell.paragraphs:
                split_plus_runs(paragraph)

            for nested_table in cell.tables:
                process_table(nested_table)


def beautify_docx(docx_path):

    print(f"Processing: {os.path.basename(docx_path)}")
    logger.info("Processing document: %s", docx_path)

    doc = Document(docx_path)

    for table in doc.tables:
        process_table(table)

    for table in doc.tables:
        format_table(table)

    doc.save(docx_path)
    logger.info("Saved formatted document: %s", docx_path)


def main():

    logger.info("Scanning output folder: %s", OUTPUT_FOLDER)
    for root, _, filenames in os.walk(OUTPUT_FOLDER):
        for filename in filenames:

            if not filename.lower().endswith(".docx"):
                continue

            path = os.path.join(root, filename)

            try:
                beautify_docx(path)
            except (PackageNotFoundError, PermissionError):
                logger.exception("Could not format document: %s", path)
                raise SystemExit(
                    f"The Microsoft file could not be opened: {path}\n"
                    "It might be missing or still open. Please close this file "
                    "first. Thank you! :)"
                ) from None

    print("\nFinished beautifying contracts.")
    logger.info("Run completed successfully")


if __name__ == "__main__":
    main()