from pathlib import Path
import json
import re
from openpyxl import load_workbook

# --------------------------------------------------
# SETTINGS
# --------------------------------------------------

EXCEL_FILE = Path("catalog.xlsx")
OUTPUT_FILE = Path("catalog-data.json")

# Use None to load the first worksheet.
# Or change it to something like:
# SHEET_NAME = "Xcvrs only"
SHEET_NAME = None

# The script accepts any of these header names.
PART_NUMBER_HEADERS = {
    "part number",
    "t1nexus part number",
    "t1 pn",
    "pn",
    "part no",
    "part no.",
    "sku",
    "item number",
}

DATA_RATE_HEADERS = {
    "data rate",
    "datarate",
    "rate",
    "speed",
}

FORM_FACTOR_HEADERS = {
    "form factor",
    "form factor name",
    "formfactor",
    "form",
    "ff",
}


def normalize_header(value):
    """Normalize Excel header text for flexible matching."""
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " "),
    )


def find_column(headers, accepted_headers):
    """
    Return the 1-based Excel column number whose header matches
    one of the accepted names.
    """
    normalized_accepted = {
        normalize_header(name) for name in accepted_headers
    }

    for column_number, header in enumerate(headers, start=1):
        if normalize_header(header) in normalized_accepted:
            return column_number

    return None


def clean_cell(value):
    if value is None:
        return ""

    return str(value).strip()


def data_rate_value(rate):
    """
    Convert data rates to a numeric value so they can be sorted
    largest to smallest.

    Examples:
      1.6T   -> 1,600,000
      800G   ->   800,000
      100G   ->   100,000
      25G    ->    25,000
      1G     ->     1,000
      100M   ->       100

    For dual-rate values such as 40G/100G, the largest rate is used.
    Blank and N/A values go to the bottom.
    """
    text = clean_cell(rate).upper()

    if not text or text == "N/A":
        return -1

    matches = re.findall(
        r"(\d+(?:\.\d+)?)\s*(T|G|M|K)?",
        text
    )

    if not matches:
        return -1

    values = []

    for number_text, unit in matches:
        number = float(number_text)

        multiplier = {
            "T": 1_000_000,
            "G": 1_000,
            "M": 1,
            "K": 0.001,
            "": 1,
        }[unit]

        values.append(number * multiplier)

    return max(values)


def main():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Could not find {EXCEL_FILE}. "
            "Put catalog.xlsx in the same folder as this script."
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    if SHEET_NAME:
        if SHEET_NAME not in workbook.sheetnames:
            raise ValueError(
                f'Worksheet "{SHEET_NAME}" was not found. '
                f"Available sheets: {workbook.sheetnames}"
            )

        sheet = workbook[SHEET_NAME]
    else:
        sheet = workbook[workbook.sheetnames[0]]

    # Assume the first row contains the headers.
    headers = [
        sheet.cell(row=1, column=column).value
        for column in range(1, sheet.max_column + 1)
    ]

    part_number_column = find_column(
        headers,
        PART_NUMBER_HEADERS
    )

    data_rate_column = find_column(
        headers,
        DATA_RATE_HEADERS
    )

    form_factor_column = find_column(
        headers,
        FORM_FACTOR_HEADERS
    )

    missing = []

    if not part_number_column:
        missing.append("Part Number")

    if not data_rate_column:
        missing.append("Data Rate")

    if not form_factor_column:
        missing.append("Form Factor")

    if missing:
        raise ValueError(
            "Missing required column(s): "
            + ", ".join(missing)
            + "\nHeaders found: "
            + ", ".join(clean_cell(header) for header in headers)
        )

    catalog = []

    for row_number in range(2, sheet.max_row + 1):
        part_number = clean_cell(
            sheet.cell(
                row=row_number,
                column=part_number_column
            ).value
        )

        data_rate = clean_cell(
            sheet.cell(
                row=row_number,
                column=data_rate_column
            ).value
        )

        form_factor = clean_cell(
            sheet.cell(
                row=row_number,
                column=form_factor_column
            ).value
        )

        # Skip completely blank rows.
        if not part_number and not data_rate and not form_factor:
            continue

        # Skip rows without a part number so customers do not
        # see blank catalog entries.
        if not part_number:
            continue

        catalog.append({
            "partNumber": part_number,
            "dataRate": data_rate,
            "formFactor": form_factor,
        })

    # Sort largest data rate to smallest.
    # Python's sort is stable, so equal-rate products keep
    # their original Excel order.
    catalog.sort(
        key=lambda item: data_rate_value(item["dataRate"]),
        reverse=True,
    )

    OUTPUT_FILE.write_text(
        json.dumps(
            catalog,
            indent=2,
            ensure_ascii=False,
        )
        + "\n",
        encoding="utf-8",
    )

    print(
        f"Generated {OUTPUT_FILE} "
        f"with {len(catalog)} catalog items."
    )


if __name__ == "__main__":
    main()
