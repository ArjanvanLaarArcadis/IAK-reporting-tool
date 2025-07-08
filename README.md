# IAK Reporting Tool

This project is licensed under the GNU General Public License v3.0 (see LICENSE).

# Arcadis IAK Reporting Tools

This repository contains tools to automate various reporting aspects within the IAK (Instandhouding Advisering Kunstwerken) project. The tools are designed to streamline the generation of reports, manage data, and ensure consistency across deliverables.

## Features

### 1. DISK PI Report Automation
- Automates the generation of PI reports from DISK exports.
- Includes PDF attachments such as:
  - Bijlage 3: ORA including inspection plan.
  - Bijlage 6: Damage drawings.
  - Bijlage 9: Attention points for the manager.

### 2. Script Descriptions

#### `generate_aandachtspunten_beheerder.py`
Generates "Bijlage 9 - Aandachtspunten Beheerder" documents based on ORA (Onderhoudsrapportage) files. Processes relevant information, including attention points and associated images, and formats it into a predefined Word template. The final document is saved as both a Word file and a PDF.

#### `generate_bijlage_3.py`
Handles the generation of "Bijlage 3" documents based on ORA files. Identifies the most recent ORA file in a given directory, checks if a corresponding "Bijlage 3" file already exists, and generates a PDF if necessary.

#### `generate_hoogste_risicos.py`
Processes ORA data to identify the highest risks for a batch and generates a Word document summarizing these risks. This document is used in discussions with asset owners to address identified risks.

#### `generate_pi_rapportage.py`
Automates the processing of Excel PI Reports from DISK. Updates configuration variables and populates the PI reports with relevant data. Includes functionality to export the reports to PDF.

#### `get_voortgang.py`
Retrieves and processes the 'voortgang' data from an Excel file. Provides functions to clean and transform the data, as well as fetch specific parameters for a given BH_code.

#### `ora_to_word.py`
Processes delivery lists, extracts relevant data from Excel files, formats the data into Word documents, and embeds associated photos. Includes functionality to configure document styles and handle specific formatting requirements.

#### `export_excel_to_pdf.py`
Provides utility functions to interact with Microsoft Excel using the COM interface. Includes functionalities to check if the "PERSONAL.XLSB" workbook is open, execute macros, and export Excel files to PDF.

#### `utils.py`
Contains utility functions for configuration management, logging setup, and file handling. Includes functions to load configuration parameters, find matching codes, and manage document saving.

## Directory Structure
- `build/` - Scripts related to the build process (e.g., PowerShell, shell, Docker compose).
- `docs/` - Documentation folder.
- `infra/` - Terraform-related scripts/modules.
- `lib/` - Library files.
- `src/` - Source code folder containing the main scripts.
- `test/` - Unit tests and integration tests.

## Usage
1. Clone the repository.
2. Install the required dependencies listed in `pyproject.toml`.
3. Run the desired script from the `src/` directory.

## Installation

To set up the project on your local machine, follow these steps:

### Prerequisites
- Python 3.10 or higher
- pip (Python package manager)
- Microsoft Word and Excel (for document generation and export functionalities)

### Steps
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo/anl-am-iak.git
   cd anl-am-iak
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Ensure that the required Microsoft Office applications (Word and Excel) are installed and accessible on your system.

4. Verify the installation by running a test script:
   ```bash
   python src/generate_pi_rapportage.py --help
   ```

## Contributing

We welcome contributions to improve the tools and scripts in this repository. To contribute:

1. Fork the repository.
2. Create a new branch for your feature or bug fix:
   ```bash
   git checkout -b feature-name
   ```
3. Make your changes and commit them:
   ```bash
   git commit -m "Description of changes"
   ```
4. Push your changes to your fork:
   ```bash
   git push origin feature-name
   ```
5. Open a pull request to the main repository.

## License

This project is licensed under the GNU GENERAL PUBLIC LICENSE V3. See the `LICENSE` file for details.

## Support

If you encounter any issues or have questions, please open an issue in the repository or contact the maintainers directly.

