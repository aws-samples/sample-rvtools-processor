# RVTools Processing Tool

A Python script for processing, anonymizing, and de-anonymizing RVTools exports for AWS migration planning and assessment.

This tool protects sensitive information in RVTools exports while allowing for analysis and secure sharing. It processes Excel files containing VMware environment data for migration assessment and planning.

## Recent Updates
* Enhanced VM ID handling - uses actual VM IDs with filename suffix for duplicates
* Dynamic entity naming with meaningful identifiers instead of generic ANON_ prefix
* Intelligent IP address anonymization that preserves network relationships
* Improved support for IPv6 addresses
* Better handling of consolidated files

## Key features:
* Consolidation of multiple RVTools exports into a single file
* Smart anonymization preserving infrastructure relationships
* Network topology preservation while maintaining privacy
* De-anonymization capability using secure mapping files
* Command-line interface for easy integration into workflows
* Randomly generated ID's {XXXX}

| Original Field | Anonymized Format | Tab |
|----------------|------------------|-----|
| VM name | VM-ID or VM-ID_filename | vInfo |
| DNS Name | HOST-XXXX.anon.local | vInfo |
| Resource Pool | POOL-XXXX | vInfo |
| Path | PATH-XXXX | vInfo |
| Log directory | PATH-XXXX | vInfo |
| Snapshot directory | PATH-XXXX | vInfo |
| Suspend directory | PATH-XXXX | vInfo |
| Annotation | ANONYMIZED_TEXT | vInfo |
| Host | HOST-XXXX | vInfo |
| Name | HOST-XXXX | vHealth |
| Hosts | HOST-XXXX | vHealth |
| Snapshot Name | SNAPSHOT-XXXX | vSnapshot |
| Filename | PATH-XXXX | vSnapshot |
| | | |
| **Network Fields** | | |
| IP Address | 10.x.y.z (preserves subnets) | Multiple |
| Primary IP Address | 10.x.y.z (preserves subnets) | vInfo |
| MAC Address | Hashed format (XX:XX:XX:XX:XX:XX) | vNetwork |
| IPv4 | 10.x.y.z (preserves subnets) | vNetwork |
| IPv6 | 2001:db8:xxxx:xxxx:xxxx | vNetwork |
| Switch | NET-XXXX | vNetwork |
| vSwitch | NET-XXXX | vSwitch |
| | | |
| **Infrastructure Fields** | | |
| Folder | PATH-XXXX | vInfo |
| Datacenter | DC-XXXX | vInfo |
| Cluster | CLUSTER-XXXX | vInfo |
| VI SDK Server | VCENTER-XXXX | vInfo |
| Address | NET-XXXX | vHealth |
| URL | URL-XXXX | vHealth |
| | | |
| **Storage Fields** | | |
| Disk | DISK-XXXX | vPartition |
| Disk Name | DISK-XXXX | vMultiPath |
| Disk Path | PATH-XXXX | vDisk |
| | | |
| **Other** | | |
| Min Required EVC Mode | EVC-XXXX | vInfo |
| Internal Sort Column | SORT-XXXX | System |

## Repository Structure
* `rvtools_processor.py`: The main Python script containing consolidation, anonymization, and de-anonymization functions.

## Usage Instructions

### Installation
Prerequisites:
* Python 3.10 or higher
* pip (Python package installer)
* openpyxl
* pandas

To install the required dependencies, run:
```bash
pip install openpyxl pandas
```

### Virtual Environment Setup
To create virtual environment:
```bash
python -m venv rvtools_anon
```

To activate virtual environment:

On Windows:
```bash
rvtools_anon\Scripts\activate
```

On Linux/Mac:
```bash
source rvtools_anon/bin/activate
```

### Basic Commands
Combine multiple RVTools exports:
```bash
python rvtools_processor.py consolidate input1.xlsx input2.xlsx -o consolidated.xlsx
```

Anonymize RVTools data:
```bash
python rvtools_processor.py anonymize input.xlsx -o anonymized.xlsx
```

De-anonymize data using mapping file:
```bash
python rvtools_processor.py deanonymize anonymized.xlsx -m mapping.json -o original.xlsx
```

### Common Use Cases
Combine and anonymize in one operation:
```bash
python rvtools_processor.py both input1.xlsx input2.xlsx -o consolidated_anonymized.xlsx
```

Preview anonymization without creating files:
```bash
python rvtools_processor.py anonymize input.xlsx --dry-run
```

View help:
```bash
python rvtools_processor.py -h
```

## Additional Features

### Network Preservation
The tool now maintains network relationships while anonymizing:
* Preserves subnet relationships between VMs
* Maintains VLAN and network segregation
* Supports both IPv4 and IPv6 addressing
* Consistent MAC address anonymization

### Consolidated File Handling
* Automatic handling of duplicate VM IDs across multiple files
* Appends filename suffix to maintain uniqueness
* Preserves relationships across consolidated data

## Troubleshooting
1. Issue: Script fails to run due to missing module
   - Error message: `ModuleNotFoundError: No module named 'pandas'`
   - Solution: Install required package using `pip install pandas`

2. Issue: Invalid file format
   - Error message: `ValueError: File format not supported`
   - Solution: Ensure input files are valid .xlsx format

3. Issue: Permission errors
   - Error message: `PermissionError: [Errno 13] Permission denied`
   - Solution: Verify write permissions in output directory

## Data Flow

The data flow in this application follows these steps:

1. Input: RVTools export files (Excel format)
2. Processing:
   - Anonymization: Replace sensitive data with unique identifiers while preserving relationships
   - De-anonymization: Map identifiers back to original values
3. Output: New processed files (anonymized or de-anonymized)

## Read More

https://aws.amazon.com/blogs/TBC

## Security

See [CONTRIBUTING](CONTRIBUTING.md#security-issue-notifications) for more information.

## License

This library is licensed under the MIT-0 License. See the LICENSE file.
```
