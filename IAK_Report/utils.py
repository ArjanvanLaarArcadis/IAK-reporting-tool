"""
Utility module for IAK Report generation

This module provides utility functions for handling file paths, converting documents,
loading configurations, and logging for the IAK Report generation process.
"""

# Built-in imports
import os
import re
import json
import logging

# External imports
import docx
from docx2pdf import convert


# Default path to the configuration file
CONFIG_FILE = "data/config.json"


def load_config(config_path=CONFIG_FILE):
    """
    Load configuration parameters from a JSON file.

    Parameters:
        config_path (str): Path to the configuration JSON file.

    Returns:
        dict: Dictionary containing configuration parameters.
    """
    print(f"Loading configuration from {config_path}...")
    with open(config_path, "r") as f:
        config = json.load(f)
    print("Configuration loaded successfully.")
    return config


def get_matching_codes(folder_path):
    # Define the regex pattern for the object code
    # Two digits, a letter, a hyphen, three digits, a hyphen, and two digits
    pattern = r"^\d{2}[A-Z]-\d{3}-\d{2}$"

    # List all files in the folder
    all_files = os.listdir(folder_path)

    # Filter files matching the pattern
    matching_codes = [file for file in all_files if re.match(pattern, file)]

    return matching_codes


def get_object_paths_codes(batch_path=None):
    """
    Get the paths and codes of all objects in the given batch path.

    Parameters:
        batch_path (str): Path to the batch directory.

    Returns:
        list: List of tuples containing the path and code of each directory.
    """
    if not batch_path:
        config = load_config(CONFIG_FILE)
        batch_path = os.path.join(config["path_batch"], config["batch"])
    
    # Get all matching object codes in the batch path
    object_paths_codes = [
        (os.path.join(batch_path, code), code)            # Tuple: (path, code)
        for code in get_matching_codes(batch_path)        # Get all codes with an object code pattern
        if os.path.isdir(os.path.join(batch_path, code))  # Only include directories
    ]

    return object_paths_codes  # List of tuples (path, code)


def convert_docx_to_pdf(input_path: str, output_path: str) -> None:
    """
    Converts a .docx file to PDF using docx2pdf.

    Parameters:
        input_path (str): Path to the input .docx file.
        output_path (str): Path to save the output PDF file.

    Returns:
        None

    Logs:
        INFO: When the conversion is successful.
        ERROR: If an error occurs during the conversion.
    """
    try:
        logging.info("Starting conversion of '%s' to '%s'.", input_path, output_path)
        # Convert the .docx file to PDF
        convert(input_path, output_path)
        logging.info("Successfully converted '%s' to '%s'.", input_path, output_path)
    except Exception as e:
        logging.error("An error occurred during conversion: %s", e)
        raise


def find_pictures_for_object_path(object_path):
    """
    Finds the directory containing pictures that start with "Inspectiefotos"
    (case-insensitive and ignoring punctuation) and end with "verkleind"
    (case-insensitive and ignoring punctuation).

    Parameters:
        object_path (str): The path to the object directory to search in.

    Returns:
        str: The path to the matching pictures directory.

    Raises:
        FileNotFoundError: If no directory matching the rules is found.
    """
    logging.info("Searching for pictures directory in: %s", object_path)

    if not os.path.isdir(object_path):
        logging.error("Provided path '%s' is not a valid directory.", object_path)
        raise FileNotFoundError(
            f"Provided path '{object_path}' is not a valid directory."
        )

    # Helper function to normalize strings by removing punctuation and converting to lowercase
    def normalize_string(s):
        return re.sub(r"[\'\-]", "", s).lower()

    # Iterate through directories inside the object_path
    for dir_name in os.listdir(object_path):
        full_path = os.path.join(object_path, dir_name)
        logging.debug("Checking directory: %s", full_path)

        # Check if it is a directory and matches the naming pattern
        normalized_name = normalize_string(dir_name)
        if (
            os.path.isdir(full_path)
            and normalized_name.startswith("inspectiefotos")
            and normalized_name.endswith("verkleind")
        ):
            logging.info("Matching directory found: %s", full_path)
            return full_path

    # If no matching directory is found, raise an exception
    logging.error(
        "No directory found in '%s' that starts with 'Inspectiefotos' and ends with 'verkleind'.",
        object_path,
    )
    raise FileNotFoundError(
        f"No directory found in '{object_path}' that starts with 'Inspectiefotos' and ends with 'verkleind'."
    )


def update_config_with_voortgang(config, voortgang):
    variables = config
    for key, value in voortgang.items():
        variables[key] = value
    variables["save_loc"] = os.path.join(
        variables["path_batch"],
        variables["batch"],
        variables["object_code"],
        variables["save_dir"],
    )
    return variables


def setup_logger(log_file="app.log", log_level=logging.INFO):
    """
    Sets up a logger with both file and console handlers.

    Args:
        log_file (str): The name of the log file.
        log_level (int): The logging level (e.g., logging.INFO, logging.DEBUG).

    Returns:
        logging.Logger: Configured logger instance.
    """
    logger = logging.getLogger()
    logger.setLevel(log_level)

    # FileHandler for logging to a file
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    # StreamHandler for logging to the console
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(
        logging.Formatter(
            "%(asctime)s - %(levelname)s - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )

    # Add handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


def return_most_recent_ora(directory: str) -> str:
    """
    Find the most recently modified ORA file in the specified directory.

    This function searches for files starting with "ORA" and having extensions
    ".xlsm", ".xlsb", or ".xlsx". It then identifies the most recently modified file.

    Args:
        directory (str): The path to the directory to search in.

    Returns:
        str: The full path of the most recently modified ORA file.

    Raises:
        FileNotFoundError: If no files starting with "ORA" are found in the directory.
    """
    logging.info("Searching for ORA files in directory: %s", directory)

    # List all files in the directory
    files = os.listdir(directory)
    logging.debug("Files in directory: %s", files)

    # Filter files that start with "ORA" and have the correct extensions
    ora_files = [
        file
        for file in files
        if file.startswith("ORA") and file.endswith((".xlsm", ".xlsb", ".xlsx"))
    ]
    logging.debug("Filtered ORA files: %s", ora_files)

    if not ora_files:
        # Raise FileNotFoundError if no files with "ORA" are found
        logging.error("No files starting with 'ORA' found in directory: %s", directory)
        raise FileNotFoundError(
            f"No files starting with 'ORA' found in directory: {directory}"
        )

    # Get the full paths of the filtered files
    full_paths = [os.path.join(directory, file) for file in ora_files]
    logging.debug("Full paths of ORA files: %s", full_paths)

    # Find the most recently modified file
    most_recent_file = max(full_paths, key=os.path.getmtime)
    logging.info("Most recent ORA file found: %s", most_recent_file)

    return most_recent_file


def save_document(document: docx.Document, save_loc: str, file_name: str) -> None:
    """
    Save the Word document to the specified location.

    Parameters:
        document (docx.Document): The Word document object to be saved.
        save_loc (str): Directory path where the document will be saved.
        file_name (str): Name of the file to save.

    Raises:
        Exception: If the document fails to save due to an error.
    """
    try:
        # Ensure the directory exists
        os.makedirs(save_loc, exist_ok=True)

        # Construct the full save path
        save_path = os.path.join(save_loc, file_name)

        # Save the document
        document.save(save_path)
        logging.info("Document saved successfully at: %s", save_path)
    except Exception as e:
        logging.error("Failed to save document: %s", e)
        raise
