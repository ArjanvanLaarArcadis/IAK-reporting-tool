# IAK Reporting Tool - Main package initialization
# Copyright (C) 2024-2025 Arcadis Nederland B.V.
#
# SPDX-License-Identifier: GPL-3.0-or-later
# See LICENSE file for full license text.

"""
IAK Reporting Tool

Automated reporting tools for IAK (Instandhouding Advisering Kunstwerken) project.
Generates PI reports, attention points, and risk assessments from DISK exports.
"""

__version__ = "0.1.0"
__author__ = "Arcadis IAK Team"
__email__ = "iak-team@arcadis.com"

# Import main modules for easier access
from . import utils
from . import utilsxls
from . import get_voortgang

__all__ = [
    "utils",
    "utilsxls", 
    "get_voortgang",
    "generate_pi_rapportage",
    "generate_aandachtspunten_beheerder",
    "generate_hoogste_risicos",
    "export_excel_to_pdf",
    "ora_to_word",
]
