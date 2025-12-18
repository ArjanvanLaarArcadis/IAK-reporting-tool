"""
IAK Reporting Tool - Combine PI Report with Appendices
Copyright (C) 2024-2025 Arcadis Nederland B.V.

SPDX-License-Identifier: GPL-3.0-or-later
See LICENSE file for full license text.

This script combines the main PI report with its appendices into a single PDF document.
The script performs the following tasks:
- Loops through werkpakket and objects using the same pattern as other generate scripts
- Checks if the PI report and all required appendices are present in the output folder
- Combines them into a single PDF document with appendices inserted at appropriate locations

The appendices that are combined:
- Bijlage 3: ORA report (inspection plan)
- Bijlage 9: Aandachtspunten beheerder (attention points for manager)

Process:
1. Find the PI report PDF in the output folder
2. Locate Bijlage 3 and Bijlage 9 PDFs (uses most recent if multiple exist)
3. Merge all PDFs into a single document
4. Save as "[PI Report Name] - compleet.pdf" in the same directory

Dependencies:
- PyPDF2 or pypdf: For PDF merging operations
- src.utils: Custom utility functions for configuration and file handling

Usage:
Run the script as a standalone program to combine PI reports with appendices for all objects in the batch.
"""

# Built-in modules
import os
import logging
import datetime as dt
from typing import Optional, List, Tuple
from pypdf import PdfWriter, PdfReader

# Local imports
from . import utils

# Constants for file patterns
PI_RAPPORT_PATTERN = "pi rapport"
BIJLAGE_3_PATTERN = "bijlage 3"
BIJLAGE_6_PATTERN = "bijlage 6"
BIJLAGE_9_PATTERN = "bijlage 9"
EXCLUDE_COMPLEET = "compleet"


def find_most_recent_file(directory: str, pattern: str, exclude_pattern: str = None) -> Optional[str]:
    """
    Find the most recent file matching a pattern in the specified directory and its subdirectories.
    
    Parameters:
        directory (str): Directory to search in (recursively).
        pattern (str): Case-insensitive substring that the filename must contain.
        exclude_pattern (str): Optional pattern to exclude from results.
    
    Returns:
        str: Full path to the most recent matching file, or None if not found.
    """
    logging.debug(f"Searching recursively for files matching pattern '{pattern}' in [{directory}]")
    
    if not os.path.exists(directory):
        logging.warning(f"Directory does not exist: [{directory}]")
        return None
    
    matching_files = []
    
    # Search recursively for files matching the pattern
    for root, dirs, files in os.walk(directory):
        for filename in files:
            if pattern.lower() in filename.lower() and filename.lower().endswith('.pdf'):
                # Exclude files matching the exclude pattern
                if exclude_pattern and exclude_pattern.lower() in filename.lower():
                    continue
                file_path = os.path.join(root, filename)
                matching_files.append(file_path)
                logging.debug(f"Found matching file: [{filename}] in [{root}]")
    
    if not matching_files:
        logging.info(f"No files found matching pattern '{pattern}' in [{directory}] or its subdirectories")
        return None
    
    # Return the most recent file based on modification time
    most_recent = max(matching_files, key=os.path.getmtime)
    logging.info(f"Most recent file: [{os.path.basename(most_recent)}]")
    return most_recent


def find_last_page_with_text(pdf_path: str, search_text: str) -> Optional[int]:
    """
    Find the last page number in a PDF that contains the specified text.
    
    Parameters:
        pdf_path (str): Path to the PDF file.
        search_text (str): Text to search for (case-insensitive).
    
    Returns:
        int: Page number (0-indexed) of the last occurrence, or None if not found.
    """
    logging.debug(f"Searching for '{search_text}' in [{os.path.basename(pdf_path)}]")
    
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            last_page = None
            
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                if search_text.lower() in text.lower():
                    last_page = page_num
                    logging.debug(f"Found '{search_text}' on page {page_num + 1}")
            
            if last_page is not None:
                logging.info(f"Last occurrence of '{search_text}' on page {last_page + 1}")
            else:
                logging.warning(f"'{search_text}' not found in PDF")
            
            return last_page
    
    except Exception as e:
        logging.error(f"Error searching PDF: {e}")
        return None


def find_insertion_points(pi_report_path: str, appendices: List[Tuple[Optional[str], str, str]]) -> List[Tuple[Optional[int], str]]:
    """
    Find insertion points for appendices in the PI report.
    
    Parameters:
        pi_report_path (str): Path to the PI report PDF.
        appendices (list): List of tuples (pdf_path, search_pattern, name).
    
    Returns:
        list: List of tuples (insert_page_index, appendix_name) with insertion points found.
    """
    insertion_points = []
    
    for pdf_path, search_pattern, name in appendices:
        if pdf_path:
            insert_page = find_last_page_with_text(pi_report_path, search_pattern)
            if insert_page is None:
                logging.warning(f"'{name}' reference not found in PI report, will append at end")
            insertion_points.append((insert_page, pdf_path, name))
    
    return insertion_points


def build_merged_pdf(pi_report_path: str, insertion_points: List[Tuple[Optional[int], str, str]]) -> PdfWriter:
    """
    Build a merged PDF with appendices inserted at specified positions.
    
    Parameters:
        pi_report_path (str): Path to the main PI report PDF.
        insertion_points (list): List of tuples (insert_page, pdf_path, name).
    
    Returns:
        PdfWriter: Writer object with merged content.
    """
    writer = PdfWriter()
    
    # Add the main PI report
    writer.append(pi_report_path)
    logging.info(f"Added PI report: [{os.path.basename(pi_report_path)}]")
    
    # Separate insertions with page numbers from those without
    insertions_with_pages = [(page, path, name) for page, path, name in insertion_points if page is not None]
    insertions_without_pages = [(page, path, name) for page, path, name in insertion_points if page is None]
    
    # Sort by page number (insert from last to first to maintain correct indices)
    insertions_with_pages.sort(key=lambda x: x[0], reverse=True)
    
    # Insert appendices at their designated positions (from last to first)
    for insert_page, pdf_path, name in insertions_with_pages:
        # insert_page is 0-indexed, add 1 to insert after that page
        insert_position = insert_page + 1
        writer.merge(insert_position, pdf_path)
        logging.info(f"Inserted {name} after page {insert_page + 1}")
    
    # Append any appendices that don't have a reference in the document
    for _, pdf_path, name in insertions_without_pages:
        writer.append(pdf_path)
        logging.info(f"Appended {name} at end of document")
    
    return writer


def combine_pdfs(pi_report_path: str, bijlage_3_path: Optional[str], 
                 bijlage_6_path: Optional[str], bijlage_9_path: Optional[str], 
                 output_path: str) -> bool:
    """
    Combine the PI report with appendices into a single PDF.
    
    The appendices are inserted at the location where they are referenced in the PI report:
    - Finds the last page mentioning "Bijlage 3" and inserts Bijlage 3 after that page
    - Finds the last page mentioning "Bijlage 6" and inserts Bijlage 6 after that page
    - Finds the last page mentioning "Bijlage 9" and inserts Bijlage 9 after that page
    
    Parameters:
        pi_report_path (str): Path to the main PI report PDF.
        bijlage_3_path (str): Path to Bijlage 3 PDF (optional).
        bijlage_6_path (str): Path to Bijlage 6 PDF (optional).
        bijlage_9_path (str): Path to Bijlage 9 PDF (optional).
        output_path (str): Path where the combined PDF will be saved.
    
    Returns:
        bool: True if successful, False otherwise.
    """
    logging.info("Starting PDF combination process")
    
    try:
        # Define appendices to process
        appendices = [
            (bijlage_3_path, BIJLAGE_3_PATTERN, "Bijlage 3"),
            (bijlage_6_path, BIJLAGE_6_PATTERN, "Bijlage 6"),
            (bijlage_9_path, BIJLAGE_9_PATTERN, "Bijlage 9")
        ]
        
        # Find insertion points for all appendices
        insertion_points = find_insertion_points(pi_report_path, appendices)
        
        # Build the merged PDF
        writer = build_merged_pdf(pi_report_path, insertion_points)
        
        # Write the combined PDF
        logging.debug(f"Writing combined PDF to: [{output_path}]")
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        logging.info(f"Successfully created combined PDF: [{os.path.basename(output_path)}]")
        return True
        
    except Exception as e:
        logging.error(f"Failed to combine PDFs: {e}")
        return False


def process_object(object_path: str, object_code: str, config: dict, logger) -> bool:
    """
    Process a single object: find PI report and appendices, then combine them.
    
    Parameters:
        object_path (str): Path to the object directory.
        object_code (str): Code identifying the object.
        config (dict): Configuration dictionary.
        logger: Logger instance for logging.
    
    Returns:
        tuple: (success_status, missing_bijlage_6)
            - success_status (bool): True if processing succeeded, False otherwise
            - missing_bijlage_6 (bool): True if bijlage 6 was missing
    """
    logger.info(f"Processing object [{object_code}]")
    
    # Determine the output directory
    output_folder = config.get("output_folder", "")
    output_dir = os.path.join(object_path, output_folder)
    
    if not os.path.exists(output_dir):
        logger.info(f"Creating output directory: [{output_dir}]")
        os.makedirs(output_dir)
    
    # Find the PI report PDF (generated by this tooling)
    # Exclude files with "compleet" in the name
    pi_report_path = find_most_recent_file(object_path, PI_RAPPORT_PATTERN, exclude_pattern=EXCLUDE_COMPLEET)
    if not pi_report_path:
        logger.warning(f"PI report not found for object [{object_code}], skipping")
        return False, False
    
    # Find Bijlage 3 (ORA report)
    bijlage_3_path = find_most_recent_file(object_path, BIJLAGE_3_PATTERN)

    # Find Bijlage 6 (Inspectietekeningen)
    bijlage_6_path = find_most_recent_file(object_path, BIJLAGE_6_PATTERN)
    missing_bijlage_6 = bijlage_6_path is None
    if missing_bijlage_6:
        logger.warning(f"Bijlage 6 not found for object [{object_code}], will proceed without it")

    # Find Bijlage 9 (Aandachtspunten beheerder)
    bijlage_9_path = find_most_recent_file(object_path, BIJLAGE_9_PATTERN)
    
    # Skip if critical appendices are missing (3 and 9)
    if not bijlage_3_path or not bijlage_9_path:
        logger.warning(f"Critical appendices (3 or 9) missing for object [{object_code}], skipping")
        return False, missing_bijlage_6
    
    # Create output filename
    pi_basename = os.path.basename(pi_report_path)
    pi_name, pi_ext = os.path.splitext(pi_basename)
    output_filename = f"{pi_name} - compleet{pi_ext}"
    output_path = os.path.join(output_dir, output_filename)

    
    # Check if combined PDF already exists
    if os.path.exists(output_path):
        logger.info(f"Combined PDF already exists: [{output_filename}], skipping")
        return True, missing_bijlage_6
    
    # Combine the PDFs
    success = combine_pdfs(pi_report_path, bijlage_3_path, bijlage_6_path, bijlage_9_path, output_path)
    
    if success:
        logger.info(f"Successfully processed object [{object_code}]")
    else:
        logger.error(f"Failed to process object [{object_code}]")
    
    return success, missing_bijlage_6


def main() -> None:
    """
    Main function to orchestrate the PDF combination process for all objects in the batch.
    """
    # Generate timestamped log filename
    timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    log_filename = f"combine_pi_with_appendices_{timestamp}.log"
    
    logger = utils.setup_logger(log_filename)
    config = utils.load_config('./config.json')
    
    werkpakket = config.get('werkpakket', 'Unknown')
    logger.info(f"Starting PDF combination process for werkpakket [{werkpakket}]")
    
    failed_objects = []
    successful_objects = []
    skipped_objects = []
    missing_bijlage_6_objects = []
    
    # Loop through all objects in the batch
    for object_path, object_code in utils.get_object_paths_codes():
        try:
            success, missing_bijlage_6 = process_object(object_path, object_code, config, logger)
            
            if missing_bijlage_6:
                missing_bijlage_6_objects.append(object_code)
            
            if success:
                successful_objects.append(object_code)
            else:
                skipped_objects.append(object_code)
                
        except Exception as e:
            logger.error(f"Unexpected error processing object [{object_code}]: {e}")
            failed_objects.append(object_code)
    
    # Summary
    logger.info("=" * 60)
    logger.info("PDF Combination Process Summary")
    logger.info("=" * 60)
    logger.info(f"Successfully combined: {len(successful_objects)} objects")
    logger.info(f"Skipped: {len(skipped_objects)} objects")
    logger.info(f"Failed: {len(failed_objects)} objects")
    logger.info(f"Missing Bijlage 6: {len(missing_bijlage_6_objects)} objects")
    
    if successful_objects:
        logger.info(f"Successful objects: {', '.join(successful_objects)}")
    
    if skipped_objects:
        logger.warning(f"Skipped objects: {', '.join(skipped_objects)}")
    
    if missing_bijlage_6_objects:
        logger.info(f"Objects missing Bijlage 6: {', '.join(missing_bijlage_6_objects)}")
    
    if failed_objects:
        logger.error(f"Failed objects: {', '.join(failed_objects)}")
    else:
        logger.info("All objects processed successfully!")


if __name__ == "__main__":
    main()

