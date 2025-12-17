# IAK Reporting Tool

This project is licensed under the GNU General Public License v3.0 (see LICENSE).

## Overview

This repository contains tools to automate various reporting aspects within the IAK (Instandhouding Advisering Kunstwerken) project. The tools are designed to streamline the generation of reports, manage data, and ensure consistency across deliverables from DISK exports.

## Features

### 1. DISK PI Report Automation
- Automates the generation of PI reports from DISK exports
- Includes PDF export functionality with fallback mechanisms
- Supports batch processing of multiple objects

### 2. Report Generation Components
- **Bijlage 3**: ORA including inspection plan
- **Bijlage 6**: Damage drawings  
- **Bijlage 9**: Attention points for the manager
- **PI report**: PI report using DISK Excel format


## Script Descriptions

### Core Processing Scripts

#### `generate_pi_rapportage.py`
**Main PI Report Generator**
- Automates processing of Excel PI Reports from DISK exports
- Updates configuration variables and populates reports with relevant data
- Includes intelligent footer handling, image processing, and sheet formatting
- Processes multiple objects in batch mode

#### `generate_aandachtspunten_beheerder.py`
- Generates "Bijlage 9 - Aandachtspunten Beheerder" documents based on ORA files
- Processes relevant information including attention points and associated images
- Formats content into predefined Word templates
- Outputs both Word and PDF formats

#### `generate_bijlage_3.py`
- Handles generation of "Bijlage 3" documents from ORA files
- Identifies most recent ORA files in directories
- Checks for existing documents and generates PDFs when necessary

#### `generate_hoogste_risicos.py`
- Processes ORA data to identify highest risks for batches
- Generates Word documents summarizing critical risks
- Supports asset owner discussions and risk management workflows

#### `combine_pi_with_appendices.py`
**PI Report Combination Tool**
- Combines PI reports with their appendices into a single PDF document
- Automatically locates PI report, Bijlage 3, and Bijlage 9 PDFs
- Creates complete reports saved as "[PI Report Name] - compleet.pdf"
- Processes all objects in batch mode using the same loop structure as other generators

### Supporting Modules

#### `get_voortgang.py`
Retrieves and processes the 'voortgang' data from an Excel file. Provides functions to clean and transform the data, as well as fetch specific parameters for a given BH_code.

#### `ora_to_word.py`
Processes delivery lists, extracts relevant data from Excel files, formats the data into Word documents, and embeds associated photos. Includes functionality to configure document styles and handle specific formatting requirements.

#### `export_excel_to_pdf.py`
**PDF Export Utility**
- Provides Microsoft Excel COM interface interactions
- Includes fallback mechanisms for different environments

#### `utils.py` & `utilsxls.py`
**Utility Libraries**
- Configuration management and logging setup
- File handling and path resolution utilities
- Excel workbook operations and image processing
- Document saving and finalization functions

## Directory Structure

```
IAK-reporting-tool/
├── IAK_Report/              # Main source code directory
│   ├── generate_*.py        # Report generation scripts
│   ├── utils.py            # General utilities
│   ├── utilsxls.py         # Excel-specific utilities
│   └── __pycache__/        # Python cache files
├── templates/               # Word document templates
├── data/                   # Data directory (not in repo)
├── config.json.example     # Configuration template
├── pyproject.toml          # Poetry project configuration
├── requirements.txt        # Pip requirements (fallback)
├── uv.lock                 # UV package manager lock file
├── poetry.lock            # Poetry lock file
└── README.md              # This file
```

## Installation

### Test files


### Prerequisites
- **Python 3.12 or higher** (recommended)
- **Microsoft Office** (Word and Excel) for document generation and COM interface
- **Windows OS** (required for Excel COM automation)

### Package Managers
This project supports multiple package managers:
- **Poetry** (recommended) - `poetry.lock`
- **UV** (fast alternative) - `uv.lock` 
- **Pip** (fallback) - `requirements.txt`

### Setup Instructions

1. **Clone the repository:**
   ```bash
   git clone https://github.com/ArjanvanLaarArcadis/IAK-reporting-tool.git
   cd IAK-reporting-tool
   ```

2. **Choose your package manager:**

   **Option A: Poetry (Recommended)**
   ```bash
   # Install Poetry if not already installed
   pip install poetry
   
   # Install dependencies
   poetry install
   
   # Activate virtual environment
   poetry shell
   ```

   **Option B: UV (Fast)**
   ```bash
   # Install UV if not already installed
   pip install uv
   
   # Install dependencies
   uv pip install -r requirements.txt
   ```

   **Option C: Pip (Traditional)**
   ```bash
   # Create virtual environment
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   
   # Install dependencies
   pip install -r requirements.txt
   ```

3. **Configure the application:**
   ```bash
   # Copy and customize configuration
   copy config.json.example config.json
   # Edit config.json with your specific settings
   ```

4. **Set up data directory:**
   - Create a `data/` directory in your project root
   - Organize your werkpakket data in subdirectories
   - Example: `data/WP-LC-WB-24-102/object-folders/`

5. **Verify installation:**
   ```bash
   # Test the main script
   python IAK_Report/generate_pi_rapportage.py
   
   # Test Excel COM interface (if needed)
   python test_excel_com.py
   ```

## Configuration

### config.json Setup
The application requires a `config.json` file with the structure following `config.json.example`.

### Data Structure
Your data directory should follow this structure:
```
data/
└── WP-LC-WB-24-102/           # Werkpakket folder
    ├── 30F-310-01/            # Object folder
    │   ├── inspectieRapport*.xlsx
    │   ├── ORA*.xlsb
    │   └── inspectiefotos/
    └── 30G-001-05/            # Another object folder
        ├── inspectieRapport*.xlsx
        └── ...
```

## Usage

### Basic Usage
```bash
# Generate PI reports for all objects in werkpakket
python IAK_Report/generate_pi_rapportage.py

# Generate bijlage 3 - ORA reports
python IAK_Report/generate_bijlage_3.py

# Generate bijlage 9 - aandachtspunten beheerder
python IAK_Report/generate_aandachtspunten_beheerder.py

# Generate highest risk reports
python IAK_Report/generate_hoogste_risicos.py

# Combine PI reports with appendices into complete PDFs
python IAK_Report/combine_pi_with_appendices.py
```

### Advanced Usage
```bash
# Process specific object (modify config.json)
python IAK_Report/generate_pi_rapportage.py

# Combine reports using Poetry script
poetry run iak-combine-reports

# Full workflow: Generate all reports then combine them
python IAK_Report/generate_pi_rapportage.py
python IAK_Report/generate_bijlage_3.py
python IAK_Report/generate_aandachtspunten_beheerder.py
python IAK_Report/combine_pi_with_appendices.py
```

### Common Issues
- **Missing config.json**: Copy from `config.json.example` and customize
- **Path issues**: Use absolute paths in configuration
- **Permission errors**: Run as Administrator if needed
- **Missing data**: Ensure data directory structure matches expectations

If you are running into issues, please log them in the **issues tab on GitHub**. This way we can tackle the issues in a centralized manner.

## Contributing

We welcome contributions to improve the tools and scripts in this repository. To contribute:

1. **Fork the repository**
2. **Create a feature branch:**
   ```bash
   git checkout -b your-feature-name
   ```
3. **Make your changes and test them:**
   ```bash
   # Run tests (as example)
   python test_styling_excel.py
   
   # Test main functionality
   python IAK_Report/generate_pi_rapportage.py
   ```
4. **Commit your changes:**
   ```bash
   git commit -m "Description of changes"
   ```
5. **Push to your fork:**
   ```bash
   git push origin feature-name
   ```
6. **Open a pull request** to the main repository

## License

This project is licensed under the **GNU General Public License v3.0**. See the `LICENSE` file for details.


## Support

If you encounter any issues or have questions:

1. **Check the troubleshooting section** above
2. **Review existing issues** in the repository
3. **Open a new issue** with:
   - Detailed error description
   - Your environment (Python version, OS, Office version)
   - Steps to reproduce the problem
   - Log files if available

## Acknowledgments

- Built for Rijkswaterstaat IAK project
- Developed by Arcadis team
- Uses Microsoft Office COM automation for document processing

