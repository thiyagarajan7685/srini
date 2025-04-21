import os
import argparse
import pandas as pd
import logging
import sys
import fitz  
import pdfplumber
import re

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

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
                        df = pd.DataFrame(table[1:], columns=table[0])
                        tables_dict[table_index] = df
                        table_index += 1
                    except Exception as table_error:
                        logger.warning(f"Error parsing table on page {page_num} in {pdf_path}: {table_error}")
    except Exception as e:
        logger.error(f"Error extracting tables from {pdf_path}: {e}")
    return tables_dict

def process_config_rows(config_df: pd.DataFrame, pdf_map: dict) -> pd.DataFrame:
    required_columns = ['File Path', 'Field Name', 'Text Extraction', 'Table No', 'Row No', 'Column No']
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

        file_path = pdf_map.get(filename)
        if not file_path:
            logger.warning(f"File '{filename}' not found in provided directory.")
            continue

        extracted_value = None

        try:
            if pd.notna(table_no) and pd.notna(row_num) and pd.notna(col_num):
                tables = extract_tables_from_pdf(file_path)
                table = tables.get(int(table_no))
                if table is not None:
                    value = table.iloc[int(row_num) - 1, int(col_num) - 1]
                    extracted_value = re.search(regex_pattern, value)
                else:
                    logger.warning(f"Table {int(table_no)} not found in {file_path}")
            elif pd.notna(regex_pattern):
                text = extract_text_from_pdf(file_path)
                match = re.search(regex_pattern, text)
                if match:
                    extracted_value = match.group()
                else:
                    logger.warning(f"No regex match for {field_name} in {file_path}")
            else:
                logger.warning(f"No valid extraction rule for {field_name}")

        except Exception as e:
            logger.error(f"Failed to process row '{field_name}' from '{file_path}': {e}")

        result.append({
            "Field Name": field_name,
            "File Extracted Path": file_path,
            "Value": extracted_value
        })

    return pd.DataFrame(result)

def main():
    parser = argparse.ArgumentParser(description="Extract data from PDFs using config file.")
    parser.add_argument("-d", "--directory", help="Path to the directory to search for PDF files.")
    parser.add_argument("-c", "--config", help="Path to the Excel config file.")
    args = parser.parse_args()

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
    output_df = process_config_rows(config_df, pdf_)

    output_file = "extracted_output.xlsx"
    output_df.to_excel(output_file, index=False)
    logger.info(f"Output saved to {output_file}")

if __name__ == "__main__":
    main()


'''
Output - In excel
field name,file extracted(path) and value
'''