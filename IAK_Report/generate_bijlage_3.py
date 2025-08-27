"""
This script handles the generation of "Bijlage 3" documents, which are based on ORA
(Onderhoudsrapportage) files. It identifies the most recent ORA file in a given directory,
checks if a corresponding "Bijlage 3" file already exists, and generates a PDF if necessary.

The script performs the following steps:
1. Searches for files starting with "Bijlage 3" in the specified directory.
2. Identifies the most recently modified ORA file if no "Bijlage 3" file exists.
3. Runs a macro on the ORA file to export it as a PDF.
4. Saves the generated PDF in the appropriate directory.

Functions:
- `file_starts_with_bijlage3`: Checks if any file in a directory starts with "Bijlage 3".
- `run_macro_on_workbook`: Executes a macro to generate a PDF from an Excel workbook.

Dependencies:
- `os`: For file and directory operations.
- `logging`: For logging script activity and errors.
- `src.utils`: Custom utility functions for configuration and logging setup.
- `src.export_excel_to_pdf`: Handles macro execution for exporting Excel sheets to PDF.

Usage:
Run this script directly to generate "Bijlage 3" documents for all objects in the batch.
"""

# Built-in modules
import os
import logging
import datetime as dt

# Local imports
from . import utils
from . import utilsxls


def file_starts_with_bijlage3(directory: str) -> str | None:
    """
    Check if any file in the provided directory or its 'Sammie' subdirectory starts with "Bijlage 3".

    Args:
        directory (str): The path to the directory to search in.

    Returns:
        str | None: The full filename of the first file that starts with "Bijlage 3",
        or None if no such file exists.
    """

    # Check if any file in the (sub)directory starts with "Bijlage 3"
    for root, _, files in os.walk(directory):
        logging.debug(f"Checking directory: [{root}]")
        for file in files:
            if file.startswith("Bijlage 3"):
                full_path = os.path.join(root, file)
                logging.info(f"Found file: [{file}]")
                return full_path  # Return the full path of the first found file

    logging.info("No file starting with 'Bijlage 3' found in object directory.")
    return None


if __name__ == "__main__":
    # Generate timestamped log filename
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    log_filename = f"generate_bijlage_3_{timestamp}.log"
    
    # Set up logging and load configuration
    logger = utils.setup_logger(log_filename, logging.INFO)
    logging.info("Starting the script to generate Bijlage 3...")
    config = utils.load_config(config_path="./config.json")

    for object_path, object_code in utils.get_object_paths_codes():
        logging.info(f"Processing object path: {object_path}, object code: {object_code}")
        try:
            #bijlage_3 = file_starts_with_bijlage3(object_path)
            #if not bijlage_3:
            logging.info(f"Generating ORA for object [{object_code}]...")
            logging.info("Checking if ORA exists...")
            ora_path = utils.return_most_recent_ora(object_path)
            logging.info(f"ORA found: {ora_path}")
            # Find the relevant ora sheet name
            ora_sheetname = utilsxls.find_ora_sheet_name(ora_path)

            logging.info("Generating the PDF...")
            
            # Defining the name (with "Bijlage 3" and ".pdf")
            filename, ext = os.path.splitext(os.path.basename(ora_path))
            pdf_filename = os.path.join(object_path, f"Bijlage 3 - {filename}.pdf")
            utilsxls.export_to_pdf(ora_path, pdf_filename, sheet_name=ora_sheetname)
            logging.info(f"Successfully generated ORA for object [{object_code}].")
            #else:
            #    logging.info(f"ORA for object [{object_code}] already exists with name [{bijlage_3}].")
        except Exception as e:
            logging.error(f"An error occurred: {e}")
            logging.error("Failed to generate ORA for object [{object_code}].")
            continue  # Continue to the next object in case of an error
