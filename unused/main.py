import pandas as pd

# Excel-Datei einlesen
file_path = "data.xlsx"
data = pd.read_excel(file_path)

# Start ab der dritten Zeile und die Spalten ab M (Index 12) prüfen
for index, row in data.iloc[2:].iterrows():  # Ab der 3. Zeile (Index 2)
    # Hier wird eine Schleife über die 4 Gruppen (Spalten M bis P, T bis V, Z bis AB, AE bis AG) erstellt
    for group_index in range(4):  # Wiederholung von 4 mal
        # Index der aktuellen Gruppe: Jede Gruppe beginnt 5 Spalten nach der vorherigen
        start_col = 12 + group_index * 5  # Startspalte für jede Gruppe
        end_col = start_col + 3  # Ende der Gruppe (3 Zellen)

        group = row[start_col:end_col]

        # Überprüfen, welche Zelle den Wert 1 hat
        for col, value in group.items():
            if value == 1:
                print(f"Zeile {index+3}, Gruppe {group_index+1}: Wert 1 in Spalte {col}")
