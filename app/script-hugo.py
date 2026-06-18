import os
import sys
import pandas as pd
import tkinter as tk
from tkinter import filedialog

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from app.processing import load_dataframes, process_data, build_excel


def main():
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Sélectionnez les fichiers Excel",
        filetypes=[("Fichiers Excel", "*.xlsx")],
    )
    if not file_paths:
        print("Aucun fichier sélectionné.")
        return

    errors = []
    dataframes = load_dataframes(file_paths, on_error=lambda msg: errors.append(msg))
    for err in errors:
        print(err)
    if not dataframes:
        print("Aucun fichier valide.")
        return

    date_input = input("Entrez la date (YYYY-MM-DD) : ")
    try:
        trade_date = pd.to_datetime(date_input).date()
    except Exception:
        print("Format de date invalide.")
        return

    processed_df, error = process_data(dataframes, trade_date)
    if error:
        print(f"Erreur : {error}")
        return

    save_path = filedialog.asksaveasfilename(
        title="Enregistrez le fichier final",
        defaultextension=".xlsx",
        filetypes=[("Fichiers Excel", "*.xlsx")],
    )
    if not save_path:
        print("Aucun emplacement sélectionné.")
        return

    buf = build_excel(processed_df)
    with open(save_path, 'wb') as f:
        f.write(buf.read())
    print(f"Fichier enregistré : {save_path}")


if __name__ == "__main__":
    main()
