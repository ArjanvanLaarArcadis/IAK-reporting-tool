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
CONFIG_FILE = os.getenv("CONFIG_FILE", "./config.json")


def load_config(config_path=CONFIG_FILE):
    """
    Load configuration parameters from a JSON file.

    Parameters:
        config_path (str): Path to the configuration JSON file.
        by default, it looks for 'config.json' in the current directory.

    Returns:
        dict: Dictionary containing configuration parameters.
    """
    print(f"Loading configuration from [{config_path}]...")
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    with open(config_path, "r") as f:
        config = json.load(f)
    print("Configuration loaded successfully.")
    return config


def get_matching_codes(folder_path):
    # Define the regex pattern for the object code
    # Starting with Two digits, a letter, a hyphen, three digits, a hyphen, and two digits. 
    # Optional trailing characters.
    # Optional the prefix "ORA " can be present, but not required.
    # Example: "31A-001-32", "51B-002-03", "ORA 31A-001-32", "ORA 51B-002-03"
    pattern = r"^(ORA\s*)?\d{2}[A-Z]-\d{3}-\d{2}"

    # List all content in the folder
    logging.debug(f"scanning folder: {folder_path}")
    all_content = os.listdir(folder_path)
    
    # Filter content matching the pattern
    matching_content = [file for file in all_content if re.match(pattern, file)]

    return matching_content


def get_object_paths_codes(config_file=CONFIG_FILE):
    """
    Get the paths and codes of all objects in the given batch path.

    Parameters:
        batch_path (str): Path to the batch directory.

    Returns:
        list: List of tuples containing the path and code of each directory.
    """
    
    config = load_config(config_file)
    werkpakket_path = os.path.join(config["path_batch"], config["werkpakket"])
    
    # Validate that the batch directory exists
    if not os.path.isdir(werkpakket_path):
        logging.error(f"Werkpakket directory not found: {werkpakket_path}")
        raise FileNotFoundError(f"Werkpakket directory not found: {werkpakket_path}")
    logging.info(f"Werkpakket directory exists: {werkpakket_path}")

    object_paths_codes = []
    # Check if specific object codes are provided
    # These could be a single string ('31W-443-43') a list of strings (['31A-001-32', '51B-002-03']), or nothing at all.
    object_code = config["object_code"]
    if object_code:
        logging.info(f"Specific object codes provided: [{object_code}]")
        # Ensure a list of codes, even if a single code is provided (list of one)
        object_codes = (
            config["object_code"]
            if isinstance(config["object_code"], list)
            else [config["object_code"]]
        )

        for object_code in object_codes:
            object_path = os.path.join(werkpakket_path, object_code)
            if os.path.isdir(object_path):
                object_paths_codes.append((object_path, object_code))
                logging.info(f"Found object directory: {object_path}")
            else:
                logging.error(f"Object directory not found: {object_path}. Skipping." )
                continue
        return object_paths_codes
    
    # If no specific object codes, return all matching codes
    logging.info("No specific object codes provided, returning all matching sub-paths:")
    
    # Return all matching codes with their paths and stripped codes
    for object_subpath in get_matching_codes(werkpakket_path):
        # Extract the object code, allowing for an optional "ORA " prefix
        match = re.match(r"^(?:ORA\s*)?(\d{2}[A-Z]-\d{3}-\d{2})", object_subpath)
        if match:
            # Add the tuple with full path and code to the list
            object_fullpath = os.path.join(werkpakket_path, object_subpath)
            object_paths_codes.append((object_fullpath, match.group(1)))
            logging.info(f"   |-> {object_subpath}")

    return object_paths_codes  # List of tuples (path with code_name, code)


def convert_docx_to_pdf(input_path: str, output_path=None) -> None:
    """
    Converts a .docx file to PDF using docx2pdf.

    Parameters:
        input_path (str): Path to the input .docx file.
        output_path (str): Path to save the output PDF file.
        If None, the PDF will be saved in the same directory as the input file with the same name.

    Returns:
        None

    Logs:
        INFO: When the conversion is successful.
        ERROR: If an error occurs during the conversion.
    """
    if output_path is None:
        output_path = os.path.splitext(input_path)[0] + ".pdf"

    try:
        logging.info("Starting conversion of '%s' to '%s'.", input_path, output_path)
        # Convert the .docx file to PDF
        convert(input_path, output_path)
        logging.info("Successfully converted '%s' to '%s'.", input_path, output_path)
        return output_path
    except Exception as e:
        logging.error("An error occurred during conversion: %s", e)
        raise


def list_pictures_for_object(object_path):
    """
    Search in the object_path, and list ALL *.png and *.jpg files which are residing 
    in the object_paths or in each subdirectory. 
    It returns the full filenames of the pictures
    """
    picture_files = []
    for root, dirs, files in os.walk(object_path):
        for filename in files:
            if filename.lower().endswith((".png", ".jpg", ".jpeg")):
                picture_files.append(os.path.join(root, filename))
    return picture_files  # List of all fullfilenames of pictures


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
    raise FileNotFoundError(
        f"No directory found in '{object_path}' that starts with 'Inspectiefotos' and ends with 'verkleind'."
    )


def update_config_with_voortgang(config, voortgang):
    variables = config
    for key, value in voortgang.items():
        variables[key] = value
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

    # Walk through the directory and all subdirectories to find matching files
    ora_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.startswith("ORA") and file.endswith((".xlsm", ".xlsb", ".xlsx")):
                ora_files.append(os.path.join(root, file))
    logging.debug(f"Filtered ORA files (with full paths): {ora_files}")

    if not ora_files:
        # Raise FileNotFoundError if no files with "ORA" are found
        logging.error(f"No files starting with 'ORA' found in directory: {directory}")
        raise FileNotFoundError(
            f"No files starting with 'ORA' found in directory: {directory}"
        )

    # Get the full paths of the filtered files
    full_paths = [os.path.join(directory, file) for file in ora_files]
    logging.debug(f"Full paths of ORA files: {full_paths}")

    # Find the most recently modified file
    most_recent_file = max(full_paths, key=os.path.getmtime)
    logging.info(f"Most recent ORA file found: {most_recent_file}")

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
