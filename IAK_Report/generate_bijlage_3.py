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

import logging

from export_excel_to_pdf import run_macro_on_workbook
from utils import (
    load_config,
    get_object_paths_codes,
    setup_logger,
    return_most_recent_ora,
)
import os


def file_starts_with_bijlage3(directory: str) -> str | None:
    """
    Check if any file in the provided directory or its 'Sammie' subdirectory starts with "Bijlage 3".

    Args:
        directory (str): The path to the directory to search in.

    Returns:
        str | None: The name of the first file that starts with "Bijlage 3",
        or None if no such file exists.
    """
    # List all files in the provided directory
    files = os.listdir(directory)

    # Check if any file starts with "Bijlage 3"
    for file in files:
        if file.startswith("Bijlage 3"):
            logging.info("Found file: %s", file)
            return file

    # Check the 'Sammie' subdirectory if it exists
    sammie_path = os.path.join(directory, "Sammie")
    if os.path.exists(sammie_path) and os.path.isdir(sammie_path):
        files = os.listdir(sammie_path)
        for file in files:
            if file.startswith("Bijlage 3"):
                logging.info("Found file in Sammie: %s", file)
                return file

    logging.info("No file starting with 'Bijlage 3' found in object directory.")
    return None


if __name__ == "__main__":
    logger = setup_logger("generate_bijlage_3.log", logging.INFO)
    logger.info("Starting the script to generate Bijlage 3...")
    config = load_config()
    path_batch = os.path.join(config["path_batch"], config["batch"])
    for object_path, object_code in get_object_paths_codes(path_batch):
        try:
            bijlage_3 = file_starts_with_bijlage3(object_path)
            if not bijlage_3:
                logger.info("Generating ORA for object %s...", object_code)
                logger.info("Checking if ORA exists...")
                ora_path = return_most_recent_ora(object_path)
                logger.info("ORA found.")
                save_loc = os.path.join(object_path, config["save_dir"])
                if not os.path.exists(save_loc):
                    os.makedirs(save_loc)
                logger.info("Generating the PDF...")
                run_macro_on_workbook(ora_path, "ORA", "ExportActiveSheetToPDF")
                logger.info("Successfully generated ORA for object %s.", object_code)
            else:
                logger.info(
                    "ORA for object %s already exists with name %s.",
                    object_code,
                    bijlage_3,
                )
        except Exception as e:
            logger.error("An error occurred: %s", e)
            logger.error("Failed to generate ORA for object %s.", object_code)

# TODO: Gaat nog niet
