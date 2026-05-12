Define Studio - Spec Parser and Review Utility

Files included
- define_studio.pyw
- requirements.txt

What it does
- Opens an Excel specification file
- Reads a domains summary sheet such as Domains / Datasets
- Reads individual dataset sheets such as AE, DM, LB, MH, CE, etc.
- Shows dataset list on the left
- Shows variables for the selected dataset on the right
- Lets you search datasets and variables
- Exports parsed outputs to CSV and JSON files

Assumptions used
- In the Domains sheet:
  - Dataset / Domain column contains dataset name
  - Description / Dataset Label contains dataset label
- In domain sheets like AE, DM, LB:
  - Variable / Name contains variable name
  - Label / Variable Label contains variable label
- The tool also tries alternate column names when possible

How to run
Option 1:
- Double click define_studio.pyw

Option 2:
- Run with pythonw:
  pythonw define_studio.pyw

Option 3:
- Run with python:
  python define_studio.pyw

Dependencies
- pandas
- openpyxl

The script tries to install missing packages automatically if needed.

Notes
- This is a practical working GUI tool for reviewing specs.
- It is not a full Define-XML generator yet.
- It is designed as a usable base that can be extended later for metadata export, define.xml creation, codelists, value-level metadata, methods, and comments.
