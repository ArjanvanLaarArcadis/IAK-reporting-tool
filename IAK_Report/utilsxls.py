"""
This module contains utility functions for handling Excel files with openpyxl.

It includes functionalities to:

- Find the most recently modified ORA file in a directory.
- Find and delete references to images with `.mpo` extensions in the workbook.
"""

# Built-in imports
import os
import logging

# External imports
import openpyxl

def load_workbook(path: str) -> openpyxl.Workbook:
    """
    Load an Excel workbook, returning the workbook object.

    Parameters:
        path (str): Full path to the Excel workbook.

    Returns:
        openpyxl.Workbook: Loaded workbook object.
    """
    try:
        logging.debug(f"Loading Excel workbook from [{path}]...")
        wb = openpyxl.load_workbook(path)
        logging.debug("Workbook loaded successfully.")
        return wb
    except FileNotFoundError:
        logging.error(f"Error: The file at [{path}] was not found.")
        raise
    except Exception as e:
        logging.error(f"An unexpected error occurred while loading the workbook: {e}")
        raise



def find_mpo_references(workbook):
    """
    Find references to images with `.mpo` extensions in the workbook.
    MPO references lead to problems in saving the workbook.

    Args:
        workbook: The openpyxl workbook object.

    Returns:
        A list of tuples containing (sheet_name, image_reference).
    """
    mpo_references = []
    for sheet in workbook.worksheets:
        if hasattr(sheet, "_images"):  # Check if the sheet contains images
            for img in sheet._images:
                # Check if the image has a path attribute and contains '.mpo'
                if hasattr(img, "path") and ".mpo" in img.path:
                    mpo_references.append((sheet.title, img))
    return mpo_references


def delete_images(workbook, image_references):
    """
    Delete images from the workbook based on the provided references.
    I was unable to find a way to convert them. I also cannot find the mpo files in the underlying zip file.

    Args:
        workbook: The openpyxl workbook object.
        image_references: A list of tuples containing (sheet_name, image_reference).

    Returns:
        None
    """
    for sheet_name, img in image_references:
        # Get the sheet by name
        sheet = workbook[sheet_name]  
        if hasattr(sheet, "_images"):  # Ensure the sheet has images
            # Check if the image is in the sheet's images
            if img in sheet._images:  
                sheet._images.remove(img)  # Remove the image
                logging.info(f"Deleted image: {img.path} from sheet: {sheet_name}")


def save_and_finalize_workbook(wb: openpyxl.Workbook, variables: dict, save_dir: str) -> None:
    """
    Save and finalize the workbook.

    Parameters:
        wb (openpyxl.Workbook): Workbook to save.
        variables (dict): Dictionary of variables.
        target_dir (str): Directory to save the workbook.
    """
    object_code = variables.get("object_code", "UNKNOWN")
    filename_excel = f"PI rapport {object_code}.xlsx"
    filepath_excel = os.path.join(save_dir, filename_excel)

    # Create save directory if it doesn't exist
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)
        logging.info(f"Created directory: {save_dir}")

    # Remove 'Document map' sheet if it exists
    if "Document map" in wb.sheetnames:
        del wb["Document map"]

    # Set active sheet to 'Sheet2'
    wb.active = wb["Sheet2"]

    # Ensure all sheets are marked as selected
    for sheet in wb:
        sheet.views.sheetView[0].tabSelected = True

    # Save the workbook
    logging.debug(f"Saving workbook to {filepath_excel}...")
    wb.save(filepath_excel)
    logging.info(f"Workbook saved: {filename_excel}")
