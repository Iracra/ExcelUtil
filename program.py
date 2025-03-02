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
        self.root.geometry("500x400")

        # Label e pulsante per selezione file
        self.label = tk.Label(root, text="Seleziona i file Excel da elaborare", font=("Arial", 12))
        self.label.pack(pady=10)

        self.select_button = tk.Button(root, text="Seleziona File", command=self.select_files, font=("Arial", 12))
        self.select_button.pack(pady=10)

        # Label per visualizzare i file selezionati
        self.selected_files_label = tk.Label(root, text="Nessun file selezionato", font=("Arial", 10), fg="gray")
        self.selected_files_label.pack(pady=5)

        # Label e pulsante per selezione cartella output
        self.output_label = tk.Label(root, text="Seleziona la cartella di output", font=("Arial", 12))
        self.output_label.pack(pady=10)

        self.output_button = tk.Button(root, text="Seleziona Cartella", command=self.select_output, font=("Arial", 12))
        self.output_button.pack(pady=10)

        # Label per visualizzare la cartella di output selezionata
        self.selected_output_label = tk.Label(root, text="Nessuna cartella selezionata", font=("Arial", 10), fg="gray")
        self.selected_output_label.pack(pady=5)

        # Label e campo di input per la riga di intestazione
        self.header_label = tk.Label(root, text="Riga di intestazione (Excel):", font=("Arial", 12))
        self.header_label.pack(pady=10)

        self.header_entry = tk.Entry(root, font=("Arial", 12))
        self.header_entry.insert(0, "6")  # Valore di default (riga 6 in Excel = indice 5 in Python)
        self.header_entry.pack(pady=5)

        # Pulsante per avviare l'elaborazione
        self.process_button = tk.Button(root, text="Avvia Elaborazione", command=self.process_excel, font=("Arial", 12))
        self.process_button.pack(pady=20)

        # Variabili per i percorsi
        self.input_paths = []
        self.output_path = ""

    def select_files(self):
        self.input_paths = filedialog.askopenfilenames(
            title="Seleziona i file Excel da elaborare",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        if self.input_paths:
            self.selected_files_label.config(text=f"Selezionati {len(self.input_paths)} file.", fg="green")

    def select_output(self):
        self.output_path = filedialog.askdirectory(title="Seleziona cartella di output")
        if self.output_path:
            self.selected_output_label.config(text=f"Cartella di output: {self.output_path}", fg="green")

    def show_progress_popup(self):
        """Mostra una finestra popup con la barra di avanzamento."""
        self.popup = tk.Toplevel(self.root)
        self.popup.title("Avanzamento")
        self.popup.geometry("300x100")

        # Barra di avanzamento
        self.progress = ttk.Progressbar(self.popup, orient="horizontal", length=250, mode="determinate")
        self.progress.pack(pady=10)

        # Label per la percentuale
        self.progress_label = tk.Label(self.popup, text="0% completato", font=("Arial", 10))
        self.progress_label.pack(pady=5)

    def update_progress(self, value):
        """Aggiorna la barra di avanzamento e la label della percentuale."""
        self.progress['value'] = value
        self.progress_label.config(text=f"{int(value)}% completato")
        self.popup.update_idletasks()  # Forza l'aggiornamento della finestra popup

    def close_progress_popup(self):
        """Chiude la finestra popup di avanzamento."""
        self.popup.destroy()

    def process_excel(self):
        if not self.input_paths:
            messagebox.showerror("Errore", "Seleziona i file Excel da elaborare!")
            return
        if not self.output_path:
            messagebox.showerror("Errore", "Seleziona la cartella di output!")
            return

        try:
            # Mostra la finestra popup di avanzamento
            self.show_progress_popup()

            # Ottieni la riga di intestazione dall'input dell'utente
            header_row = int(self.header_entry.get()) - 1  # Converti da riga Excel a indice Python

            # Dizionario per accumulare i dati delle province
            province_data = {code: pd.DataFrame() for code in PROVINCE.keys()}
            unrecognized_data = pd.DataFrame()

            # Elabora i file in sequenza
            total_files = len(self.input_paths)
            for i, path in enumerate(self.input_paths):
                # Aggiorna la barra di avanzamento
                progress_value = (i + 1) / total_files * 100
                self.update_progress(progress_value)

                df = pd.read_excel(path, header=header_row)

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

            # Chiudi la finestra popup
            self.close_progress_popup()

        except ValueError:
            messagebox.showerror("Errore", "Inserisci un numero valido per la riga di intestazione!")
            self.close_progress_popup()
        except Exception as e:
            messagebox.showerror("Errore", f"Errore durante l'elaborazione: {str(e)}")
            self.close_progress_popup()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()