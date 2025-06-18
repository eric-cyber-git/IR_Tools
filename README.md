# Windows Incident Response Triage Toolkit

See the entire write up with screenshots here: 

This project provides a lightweight, modular workflow for performing differential analysis between a known-good Windows system and a potentially compromised one.

Built using Python and PowerShell, it enables rapid, low-cost triage in environments that lack full-fledged EDR or DFIR capabilities.

## Components

- `CSV_Merge.py`: Merges and highlights differences between baseline and target artifacts
- `ArtifactCollector.ps1`: PowerShell script for gathering forensic artifacts
- `CF_config.json`: Defines Excel conditional formatting for highlighting anomalies

## Prerequisites

- Python 3.10 or higher
- Install required packages: pandas, openpyxl, xlsxwriter

Set the following environment variables (or hardcode paths if preferred):

- WinIR_Baseline_Backup=C:\WinIR\Baseline
- WinIR_Case_Data=C:\WinIR\Case_Data\Target
- WinIR_Case_folder=C:\WinIR\Cases
- WinIR_Config_Folder=C:\WinIR\Configs


## Workflow

1. Run ArtifactCollector.ps1 on a clean system and move results to the baseline folder.

2. Run it again on the target system and move those results to the target folder.

3. Run CSV_Merge.py on your analysis machine.

4. Open the generated Excel file, enable filtering, and use "Filter by Color" to isolate anomalies.

## Customization

To add support for new artifacts:

Duplicate an existing block in CF_config.json

Update the sheet name (target_<artifact_name>)

Adjust the column letter in the formula to match your field of interest

## Why Use This?

- No expensive tooling
- Fast and portable
- Reduces review fatigue by up to 90%

Ideal for triage, validation, and training in constrained environments
