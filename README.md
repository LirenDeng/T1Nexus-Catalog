# T1Nexus Product Catalog

A simple customer-facing product catalog website for T1Nexus.

The website displays product information directly from an Excel catalog and automatically updates the deployed website through GitHub Actions.

## Live Website

https://lirendeng.github.io/T1Nexus-Catalog/

## Features

- Displays T1Nexus product catalog in a clean table
- Shows:
  - Part Number
  - Data Rate
  - Form Factor
- Product data is maintained in `catalog.xlsx`
- Automatically converts Excel data into JSON
- Automatically deploys updates to GitHub Pages
- Products are sorted by data rate from highest to lowest
- No manual HTML editing is required when products are added or removed

## Project Structure

```text
T1Nexus-Catalog/
│
├── index.html
├── catalog.xlsx
├── catalog-data.json
├── convert_catalog.py
├── README.md
│
└── .github/
    └── workflows/
        └── update-catalog.yml
