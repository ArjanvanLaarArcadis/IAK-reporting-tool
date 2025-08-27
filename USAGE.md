# IAK Reporting Tool - Quick Start Guide

## Installation

```bash
pip install iak-reporting-tool
```

## Setup

1. Create a configuration file:
```bash
# Download config template
curl -O https://raw.githubusercontent.com/ArjanvanLaarArcadis/IAK-reporting-tool/main/config.json.example
mv config.json.example config.json
# Edit config.json with your settings
```

2. Set up your data directory structure:
```
data/
└── [werkpakket]/
    └── [object-folders]/
        ├── inspectieRapport*.xlsx
        ├── ORA*.xlsb
        └── inspectiefotos/
```

## Usage

### Generate PI Reports
```bash
iak-generate-pi
```

### Generate Attention Points
```bash  
iak-generate-attention
```

### Generate Risk Reports
```bash
iak-generate-risks
```

### Python API
```python
from IAK_Report import generate_pi_rapportage
generate_pi_rapportage.main()
```

## Requirements
- Windows OS
- Microsoft Office (Excel & Word)
- Python 3.12+
