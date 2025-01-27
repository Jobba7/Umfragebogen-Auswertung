from openpyxl import load_workbook


def process_excel_row(file_path, sheet_name, row_number, cells_to_evaluate):
    # Excel-Datei laden
    workbook = load_workbook(file_path)
    sheet = workbook[sheet_name]

    # Bestimmte Zellen in der Zeile auslesen
    row_data = {cell: sheet[f"{cell}{row_number}"].value for cell in cells_to_evaluate}

    # Auswertung der Zellen (hier ein Beispiel)
    for cell, value in row_data.items():
        print(f"Zelle {cell}{row_number}: {value}")

    # Workbook schließen
    workbook.close()


# Beispielaufruf
file_path = "data.xlsx"
sheet_name = "Grundtypen_1.Tabelle_1.Versuch"
row_number = 3  # Zeile, die du verarbeiten möchtest
cells_to_evaluate = ["M", "N", "O"]  # Spalten, die du prüfen willst

process_excel_row(file_path, sheet_name, row_number, cells_to_evaluate)

# 12 spalten bis antwort 1
# 3 spalten für antwort 1
# 4 spalten abstand
