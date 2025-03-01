import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import constants

PROVINCE = globals()['constants'].PROVINCE

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Elaborazione File Excel")
        self.root.geometry("500x300")

        # Label e pulsante per selezione file
        self.label = tk.Label(root, text="Seleziona i file Excel da elaborare", font=("Arial", 12))
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Seleziona File", command=self.select_files, font=("Arial", 12))
        self.select_button.pack(pady=10)

        # Label e pulsante per selezione cartella output
        self.output_label = tk.Label(root, text="Seleziona la cartella di output", font=("Arial", 12))
        self.output_label.pack(pady=10)

        self.output_button = tk.Button(root, text="Seleziona Cartella", command=self.select_output, font=("Arial", 12))
        self.output_button.pack(pady=10)

        # Pulsante per avviare l'elaborazione
        self.process_button = tk.Button(root, text="Avvia Elaborazione", command=self.process_excel, font=("Arial", 12))
        self.process_button.pack(pady=20)

        # Barra di avanzamento
        self.progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Variabili per i percorsi
        self.input_paths = []
        self.output_path = ""

    def select_files(self):
        self.input_paths = filedialog.askopenfilenames(
            title="Seleziona i file Excel da elaborare",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if self.input_paths:
            messagebox.showinfo("Info", f"Selezionati {len(self.input_paths)} file.")

    def select_output(self):
        self.output_path = filedialog.askdirectory(title="Seleziona cartella di output")
        if self.output_path:
            messagebox.showinfo("Info", f"Cartella di output selezionata: {self.output_path}")

    def process_excel(self):
        if not self.input_paths or not self.output_path:
            messagebox.showerror("Errore", "Seleziona sia i file che la cartella di output!")
            return

        try:
            # Dizionario per accumulare i dati delle province
            province_data = {code: pd.DataFrame() for code in PROVINCE.keys()}
            unrecognized_data = pd.DataFrame()

            # Elabora i file in sequenza
            total_files = len(self.input_paths)
            for i, path in enumerate(self.input_paths):
                self.progress['value'] = (i + 1) / total_files * 100
                self.root.update_idletasks()

                df = pd.read_excel(path, header=5)

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
                        os.path.join(self.output_path, f"{safe_name}.xlsx"),
                        index=False,
                        engine='openpyxl'
                    )

            # Salva dati non riconosciuti
            if not unrecognized_data.empty:
                unrecognized_data.to_excel(
                    os.path.join(self.output_path, "000_Non_riconosciute.xlsx"),
                    index=False,
                    engine='openpyxl'
                )

            # Statistiche
            total_recognized = sum(len(df) for df in province_data.values())
            messagebox.showinfo("Completato", f"Elaborazione completata!\nFile salvati in: {self.output_path}\n"
                                            f"Totale righe elaborate: {total_recognized}\n"
                                            f"Righe non riconosciute: {len(unrecognized_data)}")

        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'elaborazione: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()