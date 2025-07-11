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

# Local imports
import utils
from export_excel_to_pdf import run_macro_on_workbook


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
        logging.debug("Checking directory: %s", root)
        for file in files:
            if file.startswith("Bijlage 3"):
                full_path = os.path.join(root, file)
                logging.info("Found file: [%s]", file)
                return full_path  # Return the full path of the first found file

    logging.info("No file starting with 'Bijlage 3' found in object directory.")
    return None


if __name__ == "__main__":
    # Set up logging and load configuration
    logger = utils.setup_logger("generate_bijlage_3.log", logging.INFO)
    logger.info("Starting the script to generate Bijlage 3...")
    config = utils.load_config()

    print("hopsa")

    path_batch = os.path.join(config["path_batch"], config["batch"])
    for object_path, object_code in utils.get_object_paths_codes(path_batch):
        try:
            bijlage_3 = file_starts_with_bijlage3(object_path)
            if not bijlage_3:
                logger.info("Generating ORA for object %s...", object_code)
                logger.info("Checking if ORA exists...")
                ora_path = utils.return_most_recent_ora(object_path)
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
