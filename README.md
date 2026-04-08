# Define.xml Studio v1

Define.xml Studio v1 is the backend foundation for an internal Define-XML and CDISC validation tool.

This version is intentionally focused on the first reliable layer:

- read internal SDTM and ADaM specification templates
- normalize metadata from both workbooks
- optionally scan SAS7BDAT and XPT datasets
- compare datasets against the specifications
- export a reviewable Excel workbook with issues and summaries

This is the correct first step before adding GUI-driven Define-XML generation, CDISC Library integration, NCI terminology pull, and CORE-based validation.

## Recommended repository structure

```text
define_xml_studio/
│
├── README.md
├── requirements.txt
├── app.py
│
├── config/
│   └── settings.json
│
├── gui/
│   ├── main_window.py
│   ├── tab_study_setup.py
│   ├── tab_spec_upload.py
│   ├── tab_dataset_scan.py
│   ├── tab_validation.py
│   └── tab_generate.py
│
├── core/
│   ├── models.py
│   ├── oid_builder.py
│   ├── validator_internal.py
│   ├── define_writer.py
│   ├── xml_validator.py
│   └── qc_writer.py
│
├── readers/
│   ├── spec_reader_sdtm.py
│   ├── spec_reader_adam.py
│   ├── dataset_reader.py
│   └── terminology_reader.py
│
├── connectors/
│   ├── cdisc_library_client.py
│   ├── nci_evs_client.py
│   └── core_runner.py
│
├── templates/
│   ├── sdtm_template.xlsx
│   └── adam_template.xlsx
│
├── outputs/
├── logs/
└── tests/
```

## What this v1 script does

The included `v1_spec_parser.py` script:

- reads the uploaded SDTM and ADaM template formats
- detects domain sheets and variable specifications
- reads supporting sheets such as domains, value metadata, and formats
- standardizes column names such as `Term` and `Terms`
- optionally reads `.sas7bdat` and `.xpt` datasets from a folder
- compares actual datasets to the specification
- writes an Excel report with:
  - Summary
  - Domain-level counts
  - Variable-level normalized metadata
  - Value metadata
  - Formats
  - Dataset comparison
  - Issues

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python v1_spec_parser.py \
  --sdtm-spec "TMP-BP-006 v1_ SDTM Specification Template.xlsx" \
  --adam-spec "TMP-BP-004 v1 ADaM-Analysis Dataset Specification Template.xlsx" \
  --data-dir "path/to/datasets" \
  --output "define_xml_studio_v1_report.xlsx"
```

If you do not have datasets yet, the script still works without `--data-dir`.

## Notes

- This version does **not** generate Define-XML yet.
- This version is the backend foundation that your GUI and Define-XML writer should sit on top of.
- Once the template structures are stable, the next step is to add:
  - Define-XML writer
  - GUI
  - CDISC Library API pull
  - NCI terminology import
  - CORE integration

## Internal template assumptions currently handled

### SDTM workbook
- domain metadata from `Domains`
- value metadata from `ValueMetadata`
- codelist/format-like data from `Formats`
- individual domain sheets such as `DM`, `AE`, `VS`, etc.

### ADaM workbook
- domain metadata from `Domains`
- value metadata from `Valuemetadata`
- codelist/format-like data from `Formats`
- individual domain sheets such as `ADSL`, `ADAE`, `ADLB`, etc.

## Next build steps

1. Freeze spec-template parsing
2. Add local project save/load
3. Add Define-XML object model
4. Generate base Define-XML
5. Add GUI
6. Add CDISC and NCI connectors
7. Add validation engine integration
