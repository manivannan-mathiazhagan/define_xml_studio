# Define.xml Studio

Define.xml Studio is an internal Python-based desktop tool designed to support **Define-XML generation** and **CDISC metadata validation** using study specifications, SAS datasets, XPT files, and official CDISC/NCI reference sources.

The tool provides a unified workflow for converting SDTM and ADaM study metadata into a structured, reviewable, and reusable process for generating `define.xml`, along with validation and discrepancy reporting.

---

## Purpose

This tool is designed to:

- Read internal **SDTM and ADaM specification templates**
- Scan study datasets from **SAS7BDAT** and **XPT** files
- Compare specification metadata against implemented datasets
- Support import and usage of **NCI/CDISC Controlled Terminology**
- Connect to **CDISC Library API** for standards metadata
- Perform validation using CDISC-based rules and internal checks
- Generate **Define-XML**
- Export validation and discrepancy reports

---

## Main Features

### Metadata Input
- Load SDTM specification workbook
- Load ADaM specification workbook
- Support internal company templates
- Handle:
  - Study-level metadata
  - Dataset-level metadata
  - Variable-level metadata
  - Value-level metadata
  - Codelists
  - Methods
  - Comments

### Dataset Reading
- Read `.sas7bdat` datasets
- Read `.xpt` datasets
- Extract dataset and variable structure
- Compare implementation vs specifications

### Controlled Terminology Support
- Import local NCI/CDISC terminology files
- Support future API-based terminology pull
- Validate codelist references

### Standards Integration
- Connect to CDISC Library API
- Support standard/version-aware validation

### Validation
- Specification completeness checks
- Dataset vs specification comparison
- Missing/extra dataset and variable detection
- Type, length, and label mismatch checks
- Codelist and terminology validation
- Define-XML readiness checks
- Optional integration with CDISC CORE rules engine

### Output
- Generate `define.xml`
- Export validation report (Excel)
- Export discrepancy log (CSV/Excel)
- Save project metadata for reuse

---

## Workflow

1. Load study information
2. Upload SDTM and ADaM specs
3. Select SAS/XPT dataset folder
4. Import or fetch controlled terminology
5. Review dataset vs metadata comparison
6. Run validation checks
7. Resolve issues
8. Generate `define.xml`
9. Export QC and discrepancy reports

---

## Expected Inputs

- SDTM specification Excel file
- ADaM specification Excel file
- SAS datasets (`.sas7bdat`)
- Transport datasets (`.xpt`)
- Optional:
  - Controlled terminology files
  - Methods/comments metadata
  - Value-level metadata
  - Document references

---

## Expected Outputs

- `define.xml`
- Validation report (Excel)
- Discrepancy log (CSV/Excel)
- Processing log
- Optional metadata snapshot (JSON)

---

## Technology Stack

- **PySide6** – GUI
- **pandas** – data handling
- **openpyxl** – Excel processing
- **pyreadstat** – SAS/XPT reading
- **requests** – API integration
- **lxml** – XML generation/validation
- **sqlite3** – local caching and project storage

---

## Project Scope

This tool is intended for internal use and built around company-specific SDTM and ADaM specification templates.

It is designed to support:

- Metadata review
- Define-XML generation
- Standards-aligned validation
- Controlled terminology integration
- Study-level QC workflows

---

## Important Note

Define-XML generation depends on **complete and accurate specification metadata**.

Datasets alone are not sufficient to derive:

- Origins
- Methods
- Comments
- Value-level metadata
- Controlled terminology intent
- Computational descriptions
- Document references

Specifications remain the primary source of truth.

---

## Development Scope (v1)

The initial version will focus on:

- SDTM and ADaM spec template ingestion
- SAS/XPT dataset scanning
- Metadata comparison and discrepancy reporting
- Controlled terminology import
- Basic Define-XML generation
- Internal validation workflows

---

## Future Enhancements

- Full CDISC Library API integration
- Automated NCI terminology updates
- Advanced Define-XML features (ARM, VLM enhancements)
- Project save/load workflows
- Packaging for enterprise deployment
- Enhanced validation dashboards

---

## Next Steps

To proceed with implementation:

- Review SDTM and ADaM specification templates
- Align tool with:
  - Sheet structure
  - Column names
  - Required metadata fields
  - Template-specific logic

---

## License

Internal use only.