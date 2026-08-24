Excel-Powered Customer Catalog

This setup lets you maintain the customer catalog from one Excel file.

Files

catalog.xlsx — your catalog source

convert_catalog.py — converts Excel to JSON

catalog-data.json — website data generated automatically

index.html — public customer catalog

.github/workflows/update-catalog.yml — GitHub automation

Required Excel columns

The converter looks for these three fields:

Part Number

Data Rate

Form Factor

It also accepts headers such as:

T1Nexus Part Number

T1 PN

PN

Form Factor Name

Speed

Other Excel columns are ignored.

Normal workflow

Edit catalog.xlsx.

Add, remove, or change catalog rows.

Save the Excel file.

Replace the old catalog.xlsx in your GitHub repository.

Commit and push the change.

GitHub Actions automatically runs convert_catalog.py.

catalog-data.json is regenerated.

The customer website reads the updated JSON.

You do not need to manually edit the product JavaScript.

Sorting

The generated catalog is sorted by Data Rate from largest to smallest.

Example:

1.6T
800G
400G
200G
100G
40G
25G
10G
1G
100M

Blank or non-numeric rates are placed at the bottom.

Choose a different Excel worksheet

By default, convert_catalog.py reads the first worksheet.

To use a specific tab, change:

SHEET_NAME = None

to:

SHEET_NAME = "Xcvrs only"

Test locally

Install the dependency:

pip install openpyxl

Run:

python convert_catalog.py

Start a local server:

python -m http.server 8000

Then open:

http://localhost:8000

Do not open index.html directly with file://, because browsers may block fetch() from reading the JSON file.

GitHub Pages

If the repository is published with GitHub Pages, index.html will load catalog-data.json automatically.

The customer never sees an Excel import button or admin controls.
