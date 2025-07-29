"""
This script provides utility functions to interact with Microsoft Excel using the COM interface.

It includes functionalities to:
- Check if the "PERSONAL.XLSB" workbook is open.
- Open an Excel workbook and return the application and workbook objects.
- Execute a macro stored in "PERSONAL.XLSB".
- Close the Excel application and all open workbooks.
- Combine the steps to open a workbook, execute a macro, and close Excel.

Dependencies:
- win32com.client: For interacting with Excel via COM.
- pythoncom: For initializing and uninitializing the COM library.
- logging: For logging information and errors.

Usage:
This script is designed to be used as a utility module for automating Excel tasks, such as running macros on specific workbooks.
"""

import os
import win32com.client
import logging
import pythoncom


def is_personal_xlsb_open(excel: win32com.client.Dispatch) -> win32com.client.Dispatch:
    """
    Checks if the "PERSONAL.XLSB" workbook is open in the given Excel application instance.

    Args:
        excel: An instance of the Excel application (e.g., a COM object).

    Returns:
        The workbook object if "PERSONAL.XLSB" is open; otherwise, None.
    """
    for wb in excel.Workbooks:
        if wb.Name.lower() == "personal.xlsb":
            logging.debug("The PERSONAL.XLSB workbook is open.")
            return wb
    logging.debug("The PERSONAL.XLSB workbook is not open.")
    return None


def open_excel(
    excel_path: str,
) -> tuple[win32com.client.Dispatch, win32com.client.Dispatch]:
    """
    Opens an Excel workbook and returns the application and workbook objects.

    Parameters:
        excel_path (str): Path to the Excel file.

    Returns:
        tuple[win32com.client.Dispatch, win32com.client.Dispatch]:
            A tuple containing the Excel application object and the workbook object.

    Raises:
        FileNotFoundError: If the specified Excel file does not exist.
        Exception: If an error occurs while opening the workbook.
    """
    try:
        if not os.path.exists(excel_path):
            logging.error("The Excel file %s does not exist.", excel_path)
            raise FileNotFoundError(f"The Excel file {excel_path} does not exist.")

        # Launch Excel
        logging.info("Launching Excel application.")
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        # Open the workbook
        logging.info("Opening workbook: %s", excel_path)
        workbook = excel_app.Workbooks.Open(os.path.abspath(excel_path))
        personal_workbook = is_personal_xlsb_open(excel_app)
        if personal_workbook is None:
            logging.info("Opening PERSONAL.XLSB workbook.")
            personal_workbook = excel_app.Workbooks.Open(
                r"C:\Users\knoppers1634\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"
            )
        return excel_app, workbook
    except Exception as e:
        logging.error("An error occurred while opening the workbook: %s", e)
        raise


def execute_macro(
    excel_app: win32com.client.Dispatch,
    module_name: str,
    macro_name: str,
    local_path: str = "",
) -> None:
    """
    Executes a macro stored in PERSONAL.XLSB.

    Parameters:
        excel_app (win32com.client.Dispatch): Excel application object.
        module_name (str): Name of the module in PERSONAL.XLSB where the macro is located.
        macro_name (str): Name of the macro (Sub) in PERSONAL.XLSB to call.
        local_path (str, optional): Local path to pass as an argument to the macro. Defaults to an empty string.

    Raises:
        Exception: If an error occurs while executing the macro.
    """
    try:
        # Construct the fully qualified macro name
        full_macro_name = f"PERSONAL.XLSB!{module_name}.{macro_name}"

        # Run the macro
        excel_app.Application.Run(full_macro_name, local_path)
        logging.info("Successfully ran macro %s.", full_macro_name)
    except Exception as e:
        logging.error("An error occurred while executing the macro: %s", e)
        raise


def close_excel(
    excel_app: win32com.client.Dispatch, save_changes: bool = False
) -> None:
    """
    Closes the Excel application and all open workbooks.

    Parameters:
        excel_app (win32com.client.Dispatch): Excel application object.
        save_changes (bool, optional): Whether to save changes to the workbooks. Defaults to False.

    Raises:
        Exception: If an error occurs while closing Excel.
    """
    try:
        workbooks = [wb for wb in excel_app.Workbooks]
        for wb in workbooks:
            logging.info("Closing workbook: %s", wb.Name)
            wb.Close(SaveChanges=save_changes)
        excel_app.Quit()
    except Exception as e:
        logging.error("An error occurred while closing Excel: %s", e)


def run_macro_on_workbook(excel_path: str, module_name: str, macro_name: str) -> None:
    """
    Combines the steps to open an Excel workbook, execute a macro, and close Excel.

    Parameters:
        excel_path (str): Path to the Excel file.
        module_name (str): Name of the module in PERSONAL.XLSB where the macro is located.
        macro_name (str): Name of the macro (Sub) in PERSONAL.XLSB to call.

    Raises:
        Exception: If an error occurs during the process.
    """
    excel_app = None
    workbook = None
    try:
        # Step 1: Open Excel
        logging.info("Initializing COM library and opening Excel workbook.")
        pythoncom.CoInitialize()
        excel_app, workbook = open_excel(excel_path)

        # Step 2: Execute the macro
        logging.info(
            f"Executing macro [{module_name}.{macro_name}] "
            f"on workbook [{os.path.basename(excel_path)}]."
        )
        execute_macro(
            excel_app, module_name, macro_name, local_path=os.path.dirname(excel_path)
        )
    except Exception as e:
        logging.error("An error occurred during the process: %s", e)
    finally:
        # Step 3: Close Excel without saving changes
        if excel_app and workbook:
            logging.info("Closing Excel application.")
            close_excel(excel_app, save_changes=False)
        pythoncom.CoUninitialize()
        logging.info("COM library uninitialized.")
