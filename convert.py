#!/usr/bin/env python3

import re
import sys
from pathlib import Path

import camelot
import pandas as pd


# ============================================================
# Helpers
# ============================================================

DATE_PATTERN = r"\d{2}/\d{2}/\d{4}"
AMOUNT_PATTERN = r"-?\s*RD\$\s*[\d,.]+"


def is_transaction_table(df):
    """Return True if the table contains the expected headers."""
    text = " ".join(
        str(value)
        for value in df.astype(str).values.flatten()
    )

    required_headers = [
        "Fecha",
        "Entrada",
        "Descripción",
        "Monto",
    ]

    return all(header in text for header in required_headers)


def parse_amount(value):
    """Normalize amount and invert its sign."""
    value = value.strip()

    negative = "-" in value

    value = value.replace("RD$", "")
    value = value.replace(" ", "")
    value = value.replace(",", "")
    value = re.sub(r"[^0-9.]", "", value)

    if not value:
        return None

    # Invert the sign.
    # Negative values become positive and positive values become negative.
    return value if negative else f"-{value}"


def extract_transactions(df):
    """Extract transactions from a Camelot table."""
    transactions = []

    for _, row in df.iterrows():
        values = [
            str(value).strip()
            for value in row.tolist()
            if str(value).strip()
        ]

        if not values:
            continue

        text = " ".join(values)

        # Ignore header rows
        if all(
            header in text
            for header in [
                "Fecha",
                "Entrada",
                "Descripción",
                "Monto",
            ]
        ):
            continue

        # Find dates
        dates = re.findall(DATE_PATTERN, text)

        if len(dates) < 2:
            continue

        fecha = dates[0]
        entrada = dates[1]

        # Remove the two dates
        remaining = re.sub(
            rf"^\s*{DATE_PATTERN}\s+{DATE_PATTERN}\s*",
            "",
            text,
            count=1,
        )

        # Find amount at the end of the row
        amount_match = re.search(
            rf"({AMOUNT_PATTERN})\s*$",
            remaining,
        )

        if not amount_match:
            continue

        amount_text = amount_match.group(1)

        # Everything before the amount is the description
        descripcion = remaining[
            :amount_match.start()
        ].strip()

        monto = parse_amount(amount_text)

        if monto is None:
            continue

        transactions.append(
            {
                "Fecha": fecha,
                "Entrada": entrada,
                "Descripción": descripcion,
                "Monto": monto,
            }
        )

    return pd.DataFrame(
        transactions,
        columns=[
            "Fecha",
            "Entrada",
            "Descripción",
            "Monto",
        ],
    )


# ============================================================
# Main
# ============================================================

def main():
    # --------------------------------------------------------
    # Validate arguments
    # --------------------------------------------------------

    if len(sys.argv) not in (2, 3):
        print("Usage: ./convert.py [input] [output]")
        sys.exit(1)

    input_path = Path(sys.argv[1])

    if not input_path.exists():
        print(
            f"Error: no se encontró el archivo "
            f"'{input_path}'"
        )
        sys.exit(1)

    if not input_path.is_file():
        print(
            f"Error: '{input_path}' no es un archivo."
        )
        sys.exit(1)

    if input_path.suffix.lower() != ".pdf":
        print(
            "Error: el archivo de entrada debe ser un PDF."
        )
        sys.exit(1)

    # --------------------------------------------------------
    # Determine output path
    # --------------------------------------------------------

    if len(sys.argv) == 3:
        output_path = Path(sys.argv[2])
    else:
        # By default, create the XLSX next to the script,
        # using the same filename as the input PDF.
        script_dir = Path(__file__).resolve().parent
        output_path = script_dir / f"{input_path.stem}.xlsx"

    # Create output directory if necessary
    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Extract tables
    # --------------------------------------------------------

    print(f"Leyendo: {input_path}")

    tables = camelot.read_pdf(
        str(input_path),
        pages="all",
        flavor="stream",
    )

    print(f"Tablas encontradas: {tables.n}")

    transaction_tables = []

    for index, table in enumerate(tables, start=1):
        df = table.df

        if not is_transaction_table(df):
            continue

        print(
            f"Tabla de movimientos encontrada: {index}"
        )

        transactions = extract_transactions(df)

        if not transactions.empty:
            transaction_tables.append(transactions)

    if not transaction_tables:
        print("No se encontraron movimientos.")
        sys.exit(1)

    # --------------------------------------------------------
    # Combine results
    # --------------------------------------------------------

    result = pd.concat(
        transaction_tables,
        ignore_index=True,
    )

    # Convert dates to actual Excel dates
    result["Fecha"] = pd.to_datetime(
        result["Fecha"],
        format="%d/%m/%Y",
    )

    result["Entrada"] = pd.to_datetime(
        result["Entrada"],
        format="%d/%m/%Y",
    )

    # Remove duplicates
    result = result.drop_duplicates(
        ignore_index=True,
    )

    # --------------------------------------------------------
    # Create Excel
    # --------------------------------------------------------

    with pd.ExcelWriter(
        output_path,
        engine="openpyxl",
        date_format="DD/MM/YYYY",
    ) as writer:

        result.to_excel(
            writer,
            sheet_name="Movimientos",
            index=False,
        )

        worksheet = writer.sheets["Movimientos"]

        # Column widths
        worksheet.column_dimensions["A"].width = 14
        worksheet.column_dimensions["B"].width = 14
        worksheet.column_dimensions["C"].width = 55
        worksheet.column_dimensions["D"].width = 18

        # Date formatting
        for column in ["A", "B"]:
            for cell in worksheet[column][1:]:
                cell.number_format = "DD/MM/YYYY"

        # Freeze header
        worksheet.freeze_panes = "A2"

        # Enable filters
        worksheet.auto_filter.ref = worksheet.dimensions

    # --------------------------------------------------------
    # Done
    # --------------------------------------------------------

    print()
    print(f"Excel generado: {output_path}")
    print(f"Movimientos encontrados: {len(result)}")


if __name__ == "__main__":
    main()