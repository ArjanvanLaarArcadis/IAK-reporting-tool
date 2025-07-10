import pandas as pd
import re
import logging

PATH_TO_FILE = r"data\Voortgang_IAK_test.xlsx"

cols = [
    "Batch",
    "BH_code",
    "Objectnaam",
    "Inspectietekeningen",
    "Inspecteur 1",
    "Inspecteur 2",
    "door",
    "door.1",
    "V&R-indicatie",
    "Nader onderzoek",
    "Directe maatregelen",
    r"Niet schade gerelateerde / gebruiksspecifieke risico’s",
    "Constructieve beoordeling"
]

names = {
    "TT": "Theo Test",
    "JD": "John Doe"
}

def _load_data(path_to_file=PATH_TO_FILE, columns=cols):
    # A function to load data from an Excel file using pandas.
    # It reads the specified columns from the sheet "Blad1" and skips the first row.
    data = pd.read_excel(
        path_to_file,
        engine="openpyxl",
        sheet_name="Blad1",
        skiprows=1,
        usecols=columns,
        dtype=str,
    )
    return data

def _clean_data(df):
    # Add your data cleaning steps here
    # For example, replace initials with full names
    name_cols = [
        "Inspecteur 1",
        "Inspecteur 2",
        "door",
        "door.1",
    ]
    # Replaces the abreviations in the name columns with the full names
    # and splits the names by spaces or "+"
    for col in name_cols:
        df[col] = df[col].apply(
            lambda x: ", ".join(
                names.get(
                    name.strip(), name.strip()
                )  # Apply mapping or keep the original name
                for name in re.split(r"[+\s]", str(x))  # Split on spaces or "+"
            )
        )
    return df


def get_voortgang() -> pd.DataFrame:
    """
    Retrieves and processes the 'voortgang' data.

    This function loads raw data, cleans it, and returns a processed
    pandas DataFrame containing the 'voortgang' information.

    Returns:
        pd.DataFrame: A DataFrame containing the cleaned and processed 'voortgang' data.
    """
    logging.info("Loading and cleaning voortgang data.")
    df = _load_data()
    df = _clean_data(df)
    logging.info("Data loaded and cleaned successfully.")
    return df


def get_voortgang_params(bh_code: str):
    """
    Fetches and returns a dictionary of parameters for a given BH_code from the voortgang dataset.
    This function retrieves a specific row from the voortgang dataset based on the provided BH_code.
    It validates that exactly one record matches the BH_code and extracts relevant parameters
    from the row to construct a dictionary.
    Args:
        bh_code (str): The BH_code to filter the voortgang dataset.
    Returns:
        dict: A dictionary containing the following keys and their corresponding values:
            - "opsteller" (str): The name of the person who created the record.
            - "inspecteurs" (str): A comma-separated string of inspectors.
            - "besteknummer" (str): The specification number.
            - "hulpmiddelen" (str): Tools or aids used (e.g., VKM / HM).
            - "batch" (str): The batch identifier.
            - "object_naam" (str): The name of the object.
            - "object_code" (str): The BH_code of the object.
            - "complex_code" (str): The complex code derived from the BH_code.
            - "kwaliteitsbeheerser" (str): The quality controller.
            - "venrindicatie" (str): The V&R indication.
            - "nader_onderzoek" (str): Further investigation details.
            - "directe_maatregel" (str): Immediate measures to be taken.
            - "niet_schade_gerelateerd" (str): Non-damage-related or usage-specific risks.
            - "constructieve_beoordeling" (str): Structural assessment details.
            - "inspectietekeningen" (str): Inspection drawings.
    Raises:
        ValueError: If no records are found for the given BH_code.
        ValueError: If multiple records are found for the given BH_code.
    Logs:
        - Logs information about the fetching process and any errors encountered.
        - Logs debug information for each column value retrieved.
    """
    logging.info("Fetching parameters for BH_code: %s", bh_code)
    df = get_voortgang()
    rows = df[df['BH_code'] == bh_code]

    if rows.empty:
        logging.error("No records found for BH_code: %s", bh_code)
        raise ValueError(f"No records found for BH_code: {bh_code}")
    elif len(rows) > 1:
        logging.error("Multiple records found for BH_code: %s", bh_code)
        raise ValueError(f"Multiple records found for BH_code: {bh_code}")

    row = rows.iloc[0]
    logging.info("Record found for BH_code: %s", bh_code)

    def get_value(column):
        value = row[column] if column in row and pd.notna(row[column]) else ""
        logging.debug("Value for column '%s': %s", column, value)
        return value

    result = {
        "opsteller": get_value("door"),
        "inspecteurs": ", ".join(
            [get_value("Inspecteur 1"), get_value("Inspecteur 2")]
        ),
        "besteknummer": get_value("zaakrn"),
        "hulpmiddelen": get_value("VKM / HM"),
        "batch": get_value("Batch"),
        "object_naam": get_value("Objectnaam"),
        "object_code": get_value("BH_code"),
        "complex_code": "-".join(get_value("BH_code").split("-")[:2]),
        "kwaliteitsbeheerser": get_value("door.1"),
        "venrindicatie": get_value("V&R-indicatie"),
        "nader_onderzoek": get_value("Nader onderzoek"),
        "directe_maatregel": get_value("Directe maatregelen"),
        "niet_schade_gerelateerd": get_value(
            r"Niet schade gerelateerde / gebruiksspecifieke risico’s"
        ),
        "constructieve_beoordeling": get_value("Constructieve beoordeling"),
        "inspectietekeningen": get_value("Inspectietekeningen"),
    }
    logging.info("Parameters successfully fetched for BH_code: %s", bh_code)
    return result


if __name__ == "__main__":
    df = get_voortgang()
    print(df.columns)
