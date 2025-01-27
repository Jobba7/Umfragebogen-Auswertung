import pandas as pd

# Datei laden (ersetze 'Pfad_zur_Datei.xlsx' durch den tatsächlichen Dateipfad)
file_path = "data.xlsx"
sheet_name = "Grundtypen_1.Tabelle_1.Versuch"

# Daten aus der Excel-Datei laden
data = pd.read_excel(file_path, sheet_name=sheet_name)

# Relevante Spalten auswählen: Identifikation und Antwortspalten
# Hier nehme ich an, dass die Spalten für Antworten klar benannt sind (z. B. Antwort_1_A, Antwort_1_B, ...).
# Falls die Spaltennamen komplex sind, musst du sie gegebenenfalls anpassen.
id_column = "Unnamed: 0"  # Spalte mit Teilnehmer-IDs (z. B. "Teilnehmer")
answer_columns = [col for col in data.columns if "Antwort" in col]


# Transformieren der Antworten in eine verständliche Form
def extract_responses(row):
    # Jede Gruppe von Antwortoptionen wird durch genau eine "1" gekennzeichnet
    responses = {}
    for col in answer_columns:
        if row[col] == 1:
            # Der Spaltenname gibt die Antwort an (z. B. "Antwort_1_A")
            question = "_".join(col.split("_")[:-1])  # Frage extrahieren
            option = col.split("_")[-1]  # Option extrahieren (A, B, C)
            responses[question] = option
    return responses


# Anwenden der Transformation
transformed_data = data.apply(extract_responses, axis=1)

# In ein DataFrame umwandeln
processed_data = pd.DataFrame(transformed_data.tolist(), index=data[id_column])

# Ergebnisse speichern (optional)
output_file = "Transformierte_Daten.xlsx"
processed_data.to_excel(output_file)

print(f"Die transformierten Daten wurden in {output_file} gespeichert.")
