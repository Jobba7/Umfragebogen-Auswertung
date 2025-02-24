import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl  # Stelle sicher, dass openpyxl installiert ist: `pip install openpyxl`
from preprocessing import start_preprocessing


def select_excel_file():
    """Lässt den Nutzer eine Excel-Datei auswählen."""
    file_path = filedialog.askopenfilename(
        title="Wähle eine Excel-Datei", filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")]
    )
    if file_path:
        input_file_path.set(file_path)
        messagebox.showinfo("Erfolg", "Excel-Datei erfolgreich ausgewählt.")


def process_and_save_excel():
    """Verarbeitet die ausgewählte Excel-Datei und speichert das Ergebnis."""
    input_path = input_file_path.get()
    if not input_path:
        messagebox.showerror("Fehler", "Keine Eingabedatei ausgewählt.")
        return

    # Verarbeitung

    # Dialog, um Speicherort auszuwählen
    output_path = filedialog.asksaveasfilename(
        title="Speichere die neue Excel-Datei",
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx"), ("Alle Dateien", "*.*")],
    )
    if not output_path:
        return

    try:
        # Hier kannst du deinen Verarbeitungs-Code einfügen:
        workbook = openpyxl.load_workbook(input_path)
        sheet = workbook.active

        # Beispiel: Daten aus der ersten Zelle ändern
        sheet["A1"] = "Verarbeitet"

        # Verarbeitete Datei speichern
        workbook.save(output_path)
        messagebox.showinfo("Erfolg", f"Datei erfolgreich gespeichert: {output_path}")
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
