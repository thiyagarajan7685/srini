import os
import argparse
import pandas as pd
import logging
import sys
import fitz  # PyMuPDF
import pdfplumber
import re
from difflib import SequenceMatcher

# === Logger Setup ===


def setup_logger(debug_mode: bool):
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG if debug_mode else logging.INFO)

    # Clear existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    # Always log to console
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s"))
    logger.addHandler(stream_handler)

    # If debug mode, also log to file
    if debug_mode:
        file_handler = logging.FileHandler("debug_output.log")
        file_handler.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s"))
        logger.addHandler(file_handler)

    return logger

# === Utility Functions ===


def find_heading_page(pdf_path, heading):
    """Find the page number that contains the heading."""
    doc = fitz.open(pdf_path)
    for i, page in enumerate(doc):
        text = page.get_text()
        if heading.lower() in text.lower():
            return i  # zero-based index
    return -1


def extract_all_tables(pdf_path):
    """Extract tables from all pages using pdfplumber."""
    tables_by_page = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            if tables:
                tables_by_page[i] = tables
    return tables_by_page


def get_best_matching_table(tables, page_text, heading, max_lines=10):
    """Return the table with highest similarity to the paragraph around the heading."""
    # Get paragraph around heading for matching
    context_text = "\n".join(page_text.split(heading, 1)[
                             1].strip().split('\n')[:max_lines])

    def table_to_string(table):
        return "\n".join([
            " ".join(cell if cell else "" for cell in row)  # safely join row
            for row in table if row  # skip completely empty rows
        ])

    best_score = 0
    best_table = None
    for table in tables:
        table_text = table_to_string(table)
        score = SequenceMatcher(None, context_text, table_text).ratio()
        if score > best_score:
            best_score = score
            best_table = table
    return best_table


def extract_table_by_heading(pdf_path, heading):
    page_number = find_heading_page(pdf_path, heading)
    if page_number == -1:
        print("Heading not found.")
        return None

    tables_by_page = extract_all_tables(pdf_path)

    if page_number not in tables_by_page:
        print(f"No tables found on page {page_number + 1}.")
        return None

    tables = tables_by_page[page_number]

    if len(tables) == 1:
        # assume first row is header
        return pd.DataFrame(tables[0])

    # Multiple tables: choose based on similarity to heading context
    doc = fitz.open(pdf_path)
    page_text = doc[page_number].get_text()
    best_table = get_best_matching_table(tables, page_text, heading)

    if best_table:
        best_table_ = [
            [cell.replace('\n', ' ') if cell else "" for cell in row]
            for row in best_table
        ]
        return pd.DataFrame(best_table_)
    else:
        print("No suitable table match found.")
        return None


def list_pdfs_in_directory(directory_path):
    pdf_files = []
    for root, _, files in os.walk(directory_path):
        for file in files:
            if file.lower().endswith(".pdf"):
                full_path = os.path.join(root, file)
                pdf_files.append({"filename": file, "file_path": full_path})
    return pdf_files


def extract_text_from_pdf(pdf_path: str) -> str:
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join(page.get_text("text") for page in doc)
        doc.close()
        return text.strip()
    except Exception as e:
        logger.error(f"Error extracting text from {pdf_path}: {e}")
        return ""


def extract_tables_from_pdf(pdf_path: str) -> dict[int, pd.DataFrame]:
    tables_dict = {}
    table_index = 1
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                page_tables = page.extract_tables()
                for table in page_tables:
                    try:
                        df = pd.DataFrame(table)
                        tables_dict[table_index] = df
                        table_index += 1
                    except Exception as table_error:
                        logger.warning(
                            f"Error parsing table on page {page_num} in {pdf_path}: {table_error}")
    except Exception as e:
        logger.error(f"Error extracting tables from {pdf_path}: {e}")
    return tables_dict


def table_extraction_regex(file_path, table_no, row_nums, col_nums, identifier):
    tables = extract_tables_from_pdf(file_path)
    table = None
    if identifier:
        for i in list(tables.keys()):
            tt = tables[i]
            if identifier in list(tt.iloc[0]):
                table = tt
                break
    else:
        table = tables.get(int(table_no))

    if identifier and (table is None):
        table = extract_table_by_heading(file_path, identifier)

    if (table is None) and (identifier is None):
        msg = f"Table {int(table_no)} not found in {file_path} and also no Identifier"
        logger.warning(msg)
        return None

    if ',' in str(row_nums):
        row_indices = [int(r.strip()) - 1 for r in str(row_nums).split(',')]
        logger.info(f"Multiple rows parsed: {row_indices}")
    else:
        row_indices = [int(float(row_nums)) - 1]
        logger.info(f"Single row parsed: {row_indices[0]}")

    if ',' in str(col_nums):
        col_indices = [int(c.strip()) - 1 for c in str(col_nums).split(',')]
        logger.info(f"Multiple columns parsed: {col_indices}")
    else:
        col_indices = [int(float(col_nums)) - 1]
        logger.info(f"Single column parsed: {col_indices[0]}")

    def get_cell_value(row, col):
        try:
            value = table.iloc[row, col]
            logger.info(
                f"Extracted value at (row={row+1}, col={col+1}): {value}")
            return str(value)
        except IndexError:
            logger.warning(
                f"Invalid index at (row={row+1}, col={col+1}) — skipping")
            return ""

    result = [get_cell_value(row, col)
              for row in row_indices for col in col_indices]
    return "\n".join(result)


def process_config_rows(config_df: pd.DataFrame, pdf_map: dict, text_extract: bool) -> pd.DataFrame:
    required_columns = ['File Path', 'Field Name',
                        'Text Extraction', 'Table No', 'Row No', 'Column No']
    for col in required_columns:
        if col not in config_df.columns:
            logger.error(f"Missing required column in config: {col}")
            sys.exit(1)

    result = []

    for _, row in config_df.iterrows():
        filename = row['File Path']
        field_name = row['Field Name']
        regex_pattern = row.get('Text Extraction', None)
        table_no = row.get('Table No', None)
        row_num = row.get('Row No', None)
        col_num = row.get('Column No', None)
        identifier = row.get('Identifier', None)

        file_path = pdf_map.get(filename)
        if not file_path:
            logger.warning(
                f"File '{filename}' not found in provided directory.")
            continue

        extracted_value = None
        excel_regex_pattern = pd.notna(regex_pattern)

        try:
            if (pd.notna(table_no) or identifier) and pd.notna(row_num) and pd.notna(col_num) and not excel_regex_pattern:
                extracted_value = table_extraction_regex(
                    file_path, table_no, row_num, col_num, identifier)
            elif excel_regex_pattern:
                text = extract_text_from_pdf(file_path)
                match = re.search(regex_pattern, text)
                if match:
                    extracted_value = match.group()
                    if text_extract:
                        with open(f'{filename}_output.txt', 'w') as f:
                            f.write(text)
                else:
                    logger.warning(
                        f"No regex match for {field_name} in {file_path}")
            else:
                logger.warning(f"No valid extraction rule for {field_name}")

        except Exception as e:
            logger.error(
                f"Failed to process row '{field_name}' from '{file_path}': {e}")

        result.append({
            "Field Name": field_name,
            "File Extracted Path": file_path,
            "Value": extracted_value
        })

    return pd.DataFrame(result)


def main():
    parser = argparse.ArgumentParser(
        description="Extract data from PDFs using config file.")
    parser.add_argument(
        "-t", "--text", help="Extract text and save it as a file when set to true.")
    parser.add_argument("-d", "--directory",
                        help="Path to the directory to search for PDF files.")
    parser.add_argument(
        "-c", "--config", help="Path to the Excel config file.")
    parser.add_argument("--debug", "--dg", dest="debug", action="store_true",
                        help="Enable debug mode and save logs to a file.")
    args = parser.parse_args()

    global logger
    logger = setup_logger(args.debug)

    text_extract = args.text in ['true', True]

    if not args.directory or not args.config:
        logger.error("Both --directory and --config arguments are required.")
        parser.print_help()
        sys.exit(1)

    logger.info(f"Scanning directory: {args.directory}")
    logger.info(f"Loading config file: {args.config}")

    try:
        config_df = pd.read_excel(args.config)
        logger.info("Config file loaded successfully.")
    except Exception as e:
        logger.error(f"Failed to load config file: {e}")
        sys.exit(1)

    try:
        pdf_list = list_pdfs_in_directory(args.directory)
        if not pdf_list:
            logger.warning("No PDF files found in the directory.")
        else:
            logger.info(f"Found {len(pdf_list)} PDF files.")
            pdf_df = pd.DataFrame(pdf_list)
            logger.info("\n" + pdf_df.to_string(index=False))
    except Exception as e:
        logger.error(f"Error while listing PDF files: {e}")
        sys.exit(1)

    pdf_ = {pdf['filename']: pdf['file_path'] for pdf in pdf_list}
    output_df = process_config_rows(config_df, pdf_, text_extract)

    output_file = "extracted_output.xlsx"
    output_df.to_excel(output_file, index=False)
    logger.info(f"Output saved to {output_file}")


if __name__ == "__main__":
    main()
