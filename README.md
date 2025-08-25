## RVTools Processing Tool

A Python script for processing, anonymizing, and de-anonymizing RVTools exports for AWS migration planning and assessment.

This tool protects sensitive information in RVTools exports while allowing for analysis and secure sharing. It processes CSV files containing VMware environment data for migration assessment and planning.

## Key features:
* Consolidation of multiple RVTools exports into a single file
* Anonymization of sensitive data fields with randomly generated unique identifiers
* De-anonymization capability using secure mapping files
* Command-line interface for easy integration into workflows

## Repository Structure
* `rvtools_processor.py`: The main Python script containing consolidation, anonymization, and de-anonymization functions.

## Usage Instructions

## Installation
Prerequisites:
* Python 3.10 or higher
* pip (Python package installer)
* openpyxl
* pandas

To install the required dependencies, run:
```bash
pip install openpyxl pandas
```

## Virtual Environment Setup
To activate virtual environment:

On Windows:
```bash
venv\Scripts\activate
```

On Linux/Mac:
```bash
source venv/bin/activate
```

## Basic Commands
Combine multiple RVTools exports:
```bash
python rvtools_processor.py consolidate input1.csv input2.csv -o consolidated.csv
```

Anonymize RVTools data:
```bash
python rvtools_processor.py anonymize input.csv -o anonymized.csv
```

De-anonymize data using mapping file:
```bash
python rvtools_processor.py deanonymize anonymized.csv -m mapping.json -o original.csv
```

## Common Use Cases
Combine and anonymize in one operation:
```bash
python rvtools_processor.py both input1.csv input2.csv -o consolidated_anonymized.csv
```

View help:
```bash
python rvtools_processor.py -h
```

## Troubleshooting
1. Issue: Script fails to run due to missing module
   - Error message: `ModuleNotFoundError: No module named 'pandas'`
   - Solution: Install required package using `pip install pandas`

2. Issue: Invalid file format
   - Error message: `ValueError: File format not supported`
   - Solution: Ensure input files are valid .csv or .xlsx format

3. Issue: Permission errors
   - Error message: `PermissionError: [Errno 13] Permission denied`
   - Solution: Verify write permissions in output directory

## Data Flow

The data flow in this application follows these steps:

Input: RVTools export files (CSV/xlsx format)
Processing:
Anonymization: Replace sensitive data with unique identifiers
De-anonymization: Map identifiers back to original values
Output: New processed files (anonymized or de-anonymized)

## Read More

https://aws.amazon.com/blogs/TBC

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.
