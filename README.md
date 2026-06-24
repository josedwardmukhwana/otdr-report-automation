# Fiber Report Automation Suite

A collection of Python tools designed to automate fiber optic testing and reporting workflows.

The project processes OTDR trace data, generates reports, converts workbook formats, exports PDF documentation, and automates file naming operations used during fiber network deployment and maintenance.

## Features

### Report Generation

Extracts information from SOR trace files, including:

* Fiber length
* Attenuation measurements
* Link information

The extracted data is automatically inserted into standardized reporting templates.

### Packing Report Generation

Creates structured packing reports from collected field data and trace measurements.

### Inspection Report Conversion

Transforms generated packing reports into inspection reports suitable for quality assurance and project verification workflows.

### PDF Export

Generates PDF files from Excel workbooks by exporting individual worksheets into separate PDF documents.

### Trace File Renaming

Automates renaming of SOR trace files according to predefined naming conventions, reducing manual effort and ensuring consistency across project deliverables.

## Project Structure

```text
buffering.py      - Buffering report generation
inspection.py     - Packing-to-inspection report conversion
packing.py        - Packing report creation
pdf.py            - Excel worksheet to PDF export
rename.py         - SOR trace file renaming utility
sheathing.py      - Sheathing report processing
sz.py             - SZ report processing
```

## Technologies

* Python
* Excel Automation
* PDF Generation
* Fiber Optic Testing Data Processing
* OTDR Trace Analysis

## Use Cases

* Fiber network deployment projects
* Acceptance testing documentation
* Quality assurance inspections
* Telecom infrastructure reporting
* OTDR trace management

## Requirements

Install dependencies:

```bash
pip install -r requirements.txt
```

## Disclaimer

This project was developed to automate repetitive reporting tasks within fiber optic testing and documentation workflows. Input templates and report formats may require customization for specific projects or organizations.
