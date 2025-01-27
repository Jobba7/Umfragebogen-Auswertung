# e = "external" (rot)
# ui = "unreflektiert-impulsiv" (grün)
# s = "selbstreguliert" (blau)
grid = [
    # 1. Seite
    ["s", "ui", "e"],
    ["e", "s", "ui"],
    ["ui", "e", "s"],
    ["s", "ui", "e"],
    ["e", "s", "ui"],
    ["ui", "e", "s"],
    ["s", "ui", "e"],
    # 2. Seite
    ["ui", "e", "s"],
    ["e", "s", "ui"],
    ["s", "ui", "e"],
    ["e", "ui", "s"],
    ["ui", "e", "s"],
    ["s", "e", "ui"],
    ["e", "ui", "s"],
    # 3. Seite
    ["e", "s", "ui"],
    ["s", "e", "ui"],
    ["e", "ui", "s"],
    ["ui", "s", "e"],
    ["ui", "s", "e"],
    ["ui", "e", "s"],
    ["s", "e", "ui"],
    # 4. Seite
    ["ui", "s", "e"],
    ["s", "e", "ui"],
    ["ui", "s", "e"],
    ["e", "ui", "s"],
    ["s", "e", "ui"],
    ["s", "ui", "e"],
    ["e", "ui", "s"],
]

from openpyxl import load_workbook, Workbook


def get_facette(facette, row_number):
    workbook = load_workbook("results.xlsx")
    sheet = workbook.active
    result = {"s": 0, "ui": 0, "e": 0}

    row = sheet[row_number]

    # Antworten auswerten
    for i in range(facette, len(row), 7):
        antwortABC = row[i].value
        antwortNumber = ord(antwortABC) - ord("A")
        # Antwort basierend auf dem grid auswerten
        antwort = grid[i - 1][antwortNumber]
        # Zählung in gruppe1 basierend auf der Antwort aktualisieren
        result[antwort] += 1

    return result


def get_final_result(facette):
    # facette = {"s": 0, "ui": 0, "e": 0}
    if facette["s"] == 4:
        return "Selbstreguliert"
    if facette["s"] == 3:
        return "Überwiegend selbstreguliert"

    if facette["s"] == 2:
        if facette["e"] == 1 and facette["ui"] == 1:
            return "Ansatzweise selbstreguliert"
        if facette["e"] == 2:
            return "Mischtyp selbstreguliert / external reguliert"
        if facette["ui"] == 2:
            return "Mischtyp selbstreguliert / unreflektiert-impulsiv"

    if facette["s"] == 1:
        if facette["e"] == 3:
            return "Überwiegend external reguliert"
        if facette["ui"] == 3:
            return "Überwiegend unreflektiert-impulsiv"
        if facette["e"] == 2 and facette["ui"] == 1:
            return "Ansatzweise external reguliert"
        if facette["e"] == 1 and facette["ui"] == 2:
            return "Ansatzweise unreflektiert-impulsiv"

    if facette["s"] == 0:
        if facette["e"] == 4:
            return "External reguliert"
        if facette["ui"] == 4:
            return "Unreflektiert-impulsiv"
        if facette["e"] == 2 and facette["ui"] == 2:
            return "Mischtyp external reguliert / unreflektiert-impulsiv"
        if facette["e"] == 3 and facette["ui"] == 1:
            return "Überwiegend external reguliert"
        if facette["e"] == 1 and facette["ui"] == 3:
            return "Überwiegend unreflektiert-impulsiv"

    return f"Keine Zuordnung für diesen Fall ({facette})"


for i in range(1, 7):
    print(get_facette(i, 4))
    print()
    print(get_final_result(get_facette(i, 4)))

facetten = [
    "Fähigkeit zur Einschätzung des eigenen Lernstands",
    "Fähigkeit adäquate Lernziele zu setzen",
    "Wahl einer geeigneten Lernstrategie",
    "Anwendungsgüte der Lernstrategie",
    "Fähigkeit zur Feststellung des eigenen Lernfortschritts",
    "Fähigkeit zur Anpassung des eigenen Lernens",
    "Überprüfung und Feststellung des Lernergebnisses",
]
