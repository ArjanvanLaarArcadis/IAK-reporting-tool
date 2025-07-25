"""
This script generates "Bijlage 9 - Aandachtspunten Beheerder" documents based on data
from ORA (Onderhoudsrapportage) files. It processes relevant information, including
attention points and associated images, and formats it into a predefined Word template.
The final document is saved as both a Word file and a PDF.

The script performs the following steps:
1. Loads configuration and templates.
2. Iterates through objects, retrieving their paths and codes.
3. Extracts relevant data from the ORA file.
4. Processes the data to populate the Word document with attention points.
5. Saves the document and converts it to a PDF.

Functions:
- `create_word_document`: Creates a Word document based on a template and variables.
- `extract_relevant_data`: Filters the ORA data for relevant attention points.
- `list_of_fotonummers`: Parses and processes photo numbers from the data.
- `find_foto_path`: Finds the file path for a given photo number.
- `copy_last_table`: Duplicates the last table in the Word document.
- `remove_last_table`: Removes the last table in the Word document.
- `process_aandachtspunten_beheerder`: Populates the Word document with attention points.
- `save_aandachtspunten_beheerder`: Saves the Word document to a specified location.
- `main`: Orchestrates the entire process, including error handling.

Dependencies:
- `docx`: For creating and manipulating Word documents.
- `pandas`: For handling tabular data from the ORA file.
- `os`: For file and directory operations.
- `src.utils`: Custom utility functions for configuration, image handling, and PDF conversion.
- `src.get_voortgang`: Retrieves progress parameters for objects.
- `src.ora_to_word`: Loads ORA data into a DataFrame.

Usage:
Run this script directly to generate "Bijlage 9 - Aandachtspunten Beheerder" documents
for all objects in the batch listed in config.
"""

import docx
import pandas as pd
import os
import copy
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
import time

from utils import (
    load_config,
    get_object_paths_codes,
    convert_docx_to_pdf,
    find_pictures_for_object_path,
    update_config_with_voortgang,
    return_most_recent_ora,
    setup_logger,
    save_document,
)
from get_voortgang import get_voortgang_params
from ora_to_word import load_ora
import logging


def create_word_document(template_path: str, variables: dict) -> docx.Document:
    """
    Create and configure a Word document based on a template provided by Rijkswaterstaat.

    Parameters:
        template_path (str): Path to the Word template.
        variables (dict): Dictionary containing configuration variables such as:
            - "complex_code" (str): Code of the complex.
            - "object_code" (str): Code of the object.
            - "object_naam" (str): Name of the object.

    Returns:
        docx.Document: Configured Word document.
    """
    logging.info("Creating Word document from template: %s", template_path)

    complex_code = variables.get("complex_code", "")
    object_code = variables.get("object_code", "")
    object_naam = variables.get("object_naam", "")

    logging.debug(
        "Loaded variables - complex_code: %s, object_code: %s, object_naam: %s",
        complex_code,
        object_code,
        object_naam,
    )

    document = docx.Document(template_path)
    styles = document.styles

    # Configure footer style
    if "FooterStyle" not in styles:
        logging.debug("Adding 'FooterStyle' to document styles.")
        footer_style = styles.add_style("FooterStyle", WD_STYLE_TYPE.PARAGRAPH)
        footer_style.font.underline = False
        footer_style.font.size = Pt(7)
        footer_style.font.name = "Arial"
        footer_style.font.color.rgb = RGBColor(0, 0, 0)
    else:
        logging.debug("'FooterStyle' already exists in document styles.")
        footer_style = styles["FooterStyle"]

    # Configure paragraph style
    if "Paragraph" not in styles:
        logging.debug("Adding 'Paragraph' style to document styles.")
        paragraph_style = styles.add_style("Paragraph", WD_STYLE_TYPE.PARAGRAPH)
        paragraph_style.font.underline = False
        paragraph_style.font.size = Pt(10)
        paragraph_style.font.name = "Arial"
        paragraph_style.font.color.rgb = RGBColor(0, 0, 0)
    else:
        logging.debug("'Paragraph' style already exists in document styles.")
        paragraph_style = styles["Paragraph"]

    # Configure footer content
    try:
        Footer = document.sections[0].footer
        Footer.tables[0].cell(0, 1).text = complex_code
        Footer.tables[0].cell(0, 1).paragraphs[0].style = footer_style
        Footer.tables[0].cell(1, 1).text = f"{object_code} {object_naam}"
        Footer.tables[0].cell(1, 1).paragraphs[0].style = footer_style
        logging.info("Footer content successfully configured.")
    except Exception as e:
        logging.error("Failed to configure footer content: %s", e)
        raise

    return document


def extract_relevant_data(ORA: pd.DataFrame) -> pd.DataFrame:
    """
    Filters the input DataFrame to extract rows where the column
    "Advies mutatie I-ORA & Onderhoud" contains both "aandachtspunt" and "beheerder" (case-insensitive).

    Args:
        ORA (pd.DataFrame): The input DataFrame containing the data to filter.

    Returns:
        pd.DataFrame: A filtered DataFrame containing only the rows
        where "Advies mutatie I-ORA & Onderhoud" contains both "aandachtspunt" and "beheerder".
    """
    logging.info(
        "Filtering ORA DataFrame for rows containing 'aandachtspunt' and 'beheerder' (case-insensitive)."
    )
    relevant_columns = [column for column in ORA.columns if column.startswith('Categorie')]
    select_column = relevant_columns[0] if relevant_columns else "Advies mutatie I-ORA & Onderhoud"
    return ORA[
        ORA[select_column].str.contains(
            "aandachtspunt", case=False, na=False
        )
        & ORA[select_column].str.contains(
            "beheerder", case=False, na=False
        )
    ]


def list_of_fotonummers(fotonummers: str) -> list:
    """
    Converts a cell value from a pandas Series containing photo numbers into a list of photo numbers.

    This function processes a string representation of photo numbers, which may be:
    - A comma-separated string of photo numbers.
    - A single photo number.
    - The string "nan" (interpreted as no photo numbers).

    Args:
        fotonummers (str): A cell value from a pandas Series, expected to be a string representation of photo numbers.

    Returns:
        list: A list of photo numbers as strings. Returns an empty list if the input is "nan".

    Debug Logging:
        Logs the input value and the resulting list for debugging purposes.
    """

    logging.debug("Processing fotonummers: %s", fotonummers)
    fotonummers = str(fotonummers)
    # Check if ',' exists in the string and split accordingly
    if "," in fotonummers:
        # Split the string into a list of substrings and strip each element
        fotonummers_list = [item.strip() for item in fotonummers.split(",")]
    elif fotonummers.strip() == "nan":
        logging.debug("Input is 'nan', returning an empty list.")
        return []
    else:
        # Otherwise, just create a list with a single stripped element
        fotonummers_list = [fotonummers.strip()]

    logging.debug("Resulting list of fotonummers: %s", fotonummers_list)
    return fotonummers_list


def find_foto_path(fotonummer: str, path_imgs: str) -> str:
    """
    Finds the file path of an image based on a given photo number.

    This function searches through the list of files in the specified directory
    and returns the full path of the first image file that contains the given
    photo number in its name. Both the photo number and image filenames are
    transformed to lowercase and stripped of spaces for comparison.

    Args:
        fotonummer (str): The photo number to search for in the image filenames.
        path_imgs (str): The directory path where the images are stored.

    Returns:
        str: The full file path of the matching image, or raises FileNotFoundError if no match is found.
    """
    fotonummer = fotonummer.lower().replace(" ", "")
    images = os.listdir(path_imgs)
    for image in images:
        if fotonummer in image.lower().replace(" ", ""):
            return os.path.join(path_imgs, image)

    raise FileNotFoundError(
        f"Image with fotonummer {fotonummer} not found in {path_imgs}."
    )


def copy_last_table(word_document: docx.Document) -> None:
    """
    Duplicates the last table in the Word document.

    This function takes the last table in the provided Word template document,
    creates a deep copy of it, and appends the copied table to the document.

    Args:
        word_document (docx.Document): The Word document object where the table will be duplicated.

    Returns:
        None
    """
    logging.debug("Copying the last table in the Word document.")
    template = word_document.tables[-1]
    tbl = template._tbl
    new_tbl = copy.deepcopy(tbl)
    paragraph = word_document.add_paragraph()
    paragraph._p.addnext(new_tbl)
    logging.debug("Successfully copied and appended the last table.")


def remove_last_table(word_document: docx.Document) -> None:
    """
    Removes the last table in the Word document.

    This function identifies the last table in the provided Word document
    and removes it from the document.

    Args:
        word_document (docx.Document): The Word document object from which the last table will be removed.

    Returns:
        None
    """
    logging.debug("Removing the last table in the Word document.")
    if word_document.tables:
        tbl = word_document.tables[-1]._element
        tbl.getparent().remove(tbl)
        logging.debug("Successfully removed the last table.")
    else:
        logging.warning("No tables found in the Word document to remove.")


def process_aandachtspunten_beheerder(
    word_document: docx.Document, ora_filtered: pd.DataFrame, path_imgs: str
) -> docx.Document:
    """
    Processes the aandachtspunten beheerder and populates the Word document with relevant data.

    Args:
        word_document (docx.Document): The Word document to populate.
        ora_filtered (pd.DataFrame): Filtered ORA data containing aandachtspunten.
        path_imgs (str): Path to the directory containing images.

    Returns:
        docx.Document: The updated Word document.
    """
    logging.info("Starting to process aandachtspunten beheerder.")
    cell_style = word_document.styles["Paragraph"]

    # Duplicate tables based on the number of aandachtspunten
    logging.debug("Duplicating tables for aandachtspunten.")
    if len(ora_filtered) == 1:
        remove_last_table(word_document)
        logging.info("Removed the last table for a single aandachtspunt.")
    else:
        for _ in range(len(ora_filtered) - 2):
            copy_last_table(word_document)
        logging.info("Duplicated tables for %d aandachtspunten.", len(ora_filtered) - 1)

    # Populate each table with data
    for i, (idx, row) in enumerate(ora_filtered.iterrows()):
        logging.debug("Processing row %d: %s", idx, row.to_dict())

        aandachtspunt = row["Bevinding:\n- Inspectie\n- Onderhoud\n- Overig"].partition(
            ": "
        )[0]
        bevinding_ora = row["Bevinding:\n- Inspectie\n- Onderhoud\n- Overig"].partition(
            ": "
        )[2]
        foto_column = [column for column, value in row.items() if "Foto" in column][0]
        fotos = list_of_fotonummers(row[foto_column])
        foto1 = fotos[0] if fotos else None
        foto2 = fotos[1] if len(fotos) > 1 else None
        path_foto1 = find_foto_path(foto1, path_imgs) if foto1 else None
        path_foto2 = find_foto_path(foto2, path_imgs) if foto2 else None

        logging.info(
            "Aandachtspunt: %s, Foto1: %s, Foto2: %s", aandachtspunt, foto1, foto2
        )

        word_document.tables[i].cell(0, 0).text = str("Aandachtspunt " + aandachtspunt)
        word_document.tables[i].cell(0, 0).paragraphs[0].style = cell_style
        word_document.tables[i].cell(1, 1).text = str(row["Element"].partition(",")[0])
        word_document.tables[i].cell(1, 1).paragraphs[0].style = cell_style
        word_document.tables[i].cell(2, 1).text = str(row["Bouwdeel"])
        word_document.tables[i].cell(2, 1).paragraphs[0].style = cell_style
        word_document.tables[i].cell(4, 0).text = str(bevinding_ora)
        word_document.tables[i].cell(4, 0).paragraphs[0].style = cell_style
        word_document.tables[i].cell(4, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        relevant_columns = [column for column, value in row.items() if column.startswith('Categorie')]
        select_column = relevant_columns[0] if relevant_columns else "Advies mutatie I-ORA & Onderhoud"
        word_document.tables[i].cell(6, 0).text = str(row[select_column])
        word_document.tables[i].cell(6, 0).paragraphs[0].style = cell_style
        word_document.tables[i].cell(6, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP

        if foto1:
            logging.debug("Adding Foto1 to table %d.", i)
            word_document.tables[i].cell(4, 2).paragraphs[0].add_run().add_picture(
                path_foto1, width=2350000
            )
        if foto2:
            logging.debug("Adding Foto2 to table %d.", i)
            word_document.tables[i].cell(6, 2).paragraphs[0].add_run().add_picture(
                path_foto2, width=2350000
            )
        logging.info("Processed single aandachtspunt %d.", i + 1)
    logging.info("Finished processing aandachtspunten beheerder.")
    return word_document


def save_aandachtspunten_beheerder(document: docx.Document, variables) -> str:
    """
    Save the Word document (Bijlage 9 - Aandachtspunten Beheerder) to the specified location.

    Parameters:
        document (docx.Document): The Word document object.
        variables (dict): Dictionary containing variables like 'save_loc' and 'object_code'.

    Returns:
        str: The path of the saved document.
    """
    # Extract the save location and construct the file name
    save_loc = variables["save_loc"]
    file_name = f"Bijlage 9 - Aandachtspunten Beheerder {variables['object_code']}.docx"

    # Delegate saving to the save_document function
    save_document(document, save_loc, file_name)

    # Construct and return the full save path for reference (optional)
    return os.path.join(save_loc, file_name)


def main():
    """
    Main function to orchestrate the processing of the PI report.
    """
    logger = setup_logger("generate_aandachtspunten_beheerder.log", "DEBUG")
    logger.info("Starting the generation process for aandachtspunten beheerder.")
    config_path = "./config.json"
    config = load_config(config_path=config_path)
    TEMPLATE_WORD = os.path.join(config["path_data_aandachtspunten_beheerder"], "FORMAT_Bijlage9_AandachtspuntBeheerder.docx")
    TEMPLATE_WORD_GEEN = os.path.join(config["path_data_aandachtspunten_beheerder"], "FORMAT_Bijlage9_GeenAandachtspuntBeheerder.docx")

    # Check if both template files exist
    if not os.path.exists(TEMPLATE_WORD):
        logger.error("Template file not found: %s", TEMPLATE_WORD)
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_WORD}")
    if not os.path.exists(TEMPLATE_WORD_GEEN):
        logger.error("Template file not found: %s", TEMPLATE_WORD_GEEN)
        raise FileNotFoundError(f"Template file not found: {TEMPLATE_WORD_GEEN}")

    logger.info("Template files validated successfully.")
    failed_objects = []

    for object_path, object_code in get_object_paths_codes(config_file=config_path):
        logger.info("Processing object path: %s, object code: %s", object_path, object_code)
        voortgang = get_voortgang_params(object_code)
        variables = update_config_with_voortgang(config, voortgang)
        try:
            path_ora = return_most_recent_ora(object_path)
            print("Checking for images...")
            path_imgs = find_pictures_for_object_path(object_path)
            ora = load_ora(path_ora)
            ora_filtered = extract_relevant_data(ora)
            logger.info(
                "The number of aandachtspunten voor beheerder is: %d",
                len(ora_filtered),
            )
            if len(ora_filtered) == 0:
                logger.info("Making the word document with no aandachtspunten...")
                word_document = create_word_document(TEMPLATE_WORD_GEEN, variables)
            else:
                logger.info("Making the word document with aandachtspunten...")
                word_document = create_word_document(TEMPLATE_WORD, variables)
                word_document = process_aandachtspunten_beheerder(
                    word_document, ora_filtered, path_imgs
                )
            document_path = save_aandachtspunten_beheerder(word_document, variables)
            logging.info("Word document saved successfully at: %s", document_path)
            time.sleep(1)
            convert_docx_to_pdf(
                document_path,
                os.path.join(variables["save_loc"], f"Bijlage 9 - {object_code}.pdf"),
            )
            logging.info(
                "PDF document generated successfully for object code: %s",
                object_code,
            )
            logging.info("Successfully processed object code: %s", object_code)
        except Exception as e:
            failed_objects.append(object_code)
            logging.error(
                "Failed to generate for object code: %s. Error: %s", object_code, e
            )

    logging.info(
        "Processing completed. Failed objects: [%s]",
        ", ".join(failed_objects) if failed_objects else "None",
    )

if __name__ == "__main__":
    main()
