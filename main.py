import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from analysis import evaluate_results
from preprocessing import start_preprocessing


def select_excel_file():
    """Lässt den Nutzer eine Excel-Datei auswählen."""
    file_path = filedialog.askopenfilename(
        title="Wähle eine Excel-Datei", filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")]
    )
    if file_path:
        input_file_path.set(file_path)


def process_and_save_excel():
    """Verarbeitet die ausgewählte Excel-Datei und speichert das Ergebnis."""
    input_path = input_file_path.get()
    if not input_path:
        messagebox.showerror("Fehler", "Keine Eingabedatei ausgewählt.")
        return

    # Dialog, um Speicherort auszuwählen
    output_path = filedialog.asksaveasfilename(
        title="Speichere die neue Excel-Datei",
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
    )
    if not output_path:
        return

    try:
        # Verarbeitung
        workbook = start_preprocessing(input_path)
        result = evaluate_results(workbook)

        # Verarbeitete Datei speichern
        result.save(output_path)

        # Datei mit dem Standard-Programm öffnen
        os.startfile(output_path)

        # Programm beenden
        sys.exit()

    except Exception as e:
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}")


# Hauptfenster erstellen
root = tk.Tk()
root.title("Excel-Dateiverarbeitung")

# Variable zum Speichern des Dateipfads
input_file_path = tk.StringVar()

# UI-Komponenten
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

btn_select_file = tk.Button(frame, text="Excel-Datei auswählen", command=select_excel_file)
btn_select_file.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

entry_file_path = tk.Entry(frame, textvariable=input_file_path, state="readonly", width=50)
entry_file_path.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

btn_process_save = tk.Button(frame, text="Verarbeiten und Speichern", command=process_and_save_excel)
btn_process_save.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

# Grid-Anpassungen für Fenster-Resizing
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=3)

# Hauptloop starten
root.mainloop()
