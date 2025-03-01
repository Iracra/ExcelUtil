import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os

import constants

PROVINCE = globals()['constants'].PROVINCE

def select_files():
    root = tk.Tk()
    root.withdraw()
    
    # Seleziona MULTIPLI file Excel
    input_paths = filedialog.askopenfilenames(
        title="Seleziona i file Excel da elaborare",
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )
    
    # Seleziona cartella output
    output_path = filedialog.askdirectory(title="Seleziona cartella di output")
    
    return input_paths, output_path

def process_excel():
    try:
        input_paths, output_path = select_files()
        
        # Dizionario per accumulare i dati delle province
        province_data = {code: pd.DataFrame() for code in PROVINCE.keys()}
        unrecognized_data = pd.DataFrame()

        # Elabora i file in sequenza
        for path in input_paths:
            df = pd.read_excel(path,header=5)
            
            # Trova colonna 'Provincia'
            provincia_col = next((col for col in df.columns if 'provincia' in col.lower()), None)
            if not provincia_col:
                print(f"Attenzione: colonna 'Provincia' non trovata in {os.path.basename(path)}. File saltato.")
                continue

            # Normalizza codici
            df['Codice_Provincia'] = df[provincia_col].astype(str).str.strip().str.upper()
            
            # Separa dati riconosciuti e non
            mask = df['Codice_Provincia'].isin(PROVINCE.keys())
            recognized = df[mask]
            unrecognized = df[~mask]

            # Aggiungi ai dati accumulati
            for code, group in recognized.groupby('Codice_Provincia'):
                province_data[code] = pd.concat([province_data[code], group], ignore_index=True)
            
            unrecognized_data = pd.concat([unrecognized_data, unrecognized], ignore_index=True)

        # Salva i file per provincia
        for code, df in province_data.items():
            if not df.empty:
                nome_provincia = PROVINCE[code]
                safe_name = nome_provincia.replace(" ", "_").replace("'", "")
                df.to_excel(
                    os.path.join(output_path, f"{safe_name}.xlsx"),
                    index=False,
                    engine='openpyxl'
                )

        # Salva dati non riconosciuti
        if not unrecognized_data.empty:
            unrecognized_data.to_excel(
                os.path.join(output_path, "000:non_riconosciute.xlsx"),
                index=False,
                engine='openpyxl'
            )

        # Statistiche
        total_recognized = sum(len(df) for df in province_data.values())
        print(f"Elaborazione completata! File salvati in: {output_path}")
        print(f"Totale righe elaborate: {total_recognized}")
        print(f"Righe non riconosciute: {len(unrecognized_data)}")

    except Exception as e:
        print(f"Errore durante l'elaborazione: {str(e)}")

if __name__ == "__main__":
    process_excel()