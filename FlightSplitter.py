import os
import sys
import pandas as pd
from datetime import datetime
from openpyxl.utils.datetime import from_excel
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import threading
import time

# === Conversion date Excel en texte "DDMMM" ===
def excel_date_to_dd_mmm(excel_date):
    try:
        if isinstance(excel_date, (int, float)):
            date = from_excel(excel_date)
            return date.strftime("%d%b").upper()
        elif isinstance(excel_date, (datetime, pd.Timestamp)):
            return excel_date.strftime("%d%b").upper()
        elif isinstance(excel_date, str):
            return excel_date
        else:
            return str(excel_date)
    except Exception:
        return str(excel_date)

# === Ouvrir dossier selon l'OS ===
def ouvrir_dossier(dossier):
    try:
        if sys.platform.startswith('win'):
            os.startfile(dossier)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', dossier])
        else:  # Linux
            subprocess.call(['xdg-open', dossier])
    except Exception as e:
        print(f"Impossible d'ouvrir le dossier automatiquement : {e}")

# === Fonction principale du traitement ===
def traiter_fichier(file_path, dossier_parent, progress_bar, label_progress, canvas, avion_item, root):
    try:
        if not file_path.endswith('.xlsx'):
            messagebox.showerror("Erreur", "Le fichier choisi n'est pas un .xlsx")
            return

        date_du_jour = datetime.now().strftime("%d-%m-%Y")
        dossier_sortie = os.path.join(dossier_parent, f"Rotations vols air france {date_du_jour}")
        os.makedirs(dossier_sortie, exist_ok=True)
        print(f"Dossier de sortie : {dossier_sortie}")

        df_source = pd.read_excel(file_path, header=None)
        data = df_source.iloc[4:, 1:10].copy()
        data = data.fillna("")
        identifiants_uniques = data.iloc[:, 0].dropna().unique()
        entete = ["N vol", "Date", "Dep.", "Arr.", "", "", "Dep.", "Arr."]

        total_avions = len(identifiants_uniques)
        progress_bar["maximum"] = total_avions

        canvas_width = int(canvas["width"])

        for idx, identifiant in enumerate(identifiants_uniques):
            try:
                lignes_filtrees = data[data.iloc[:, 0] == identifiant]
                valeurs_colonne_C = lignes_filtrees.iloc[:, 1].unique()
                df_sortie = pd.DataFrame(columns=entete)

                for i, (_, row) in enumerate(lignes_filtrees.iterrows()):
                    n_vol = f"AF {identifiant}"
                    date_dep = excel_date_to_dd_mmm(row.iloc[5])
                    dep = str(row.iloc[3])
                    arr = str(row.iloc[4])
                    date_arr = excel_date_to_dd_mmm(row.iloc[8])
                    dep_arr = str(row.iloc[6])
                    arr_arr = str(row.iloc[7])

                    if i > 0:
                        df_sortie.loc[len(df_sortie)] = [date_dep] + [""] * (len(entete) - 1)

                    df_sortie.loc[len(df_sortie)] = [
                        n_vol, date_dep, dep, arr, "", date_arr, dep_arr, arr_arr
                    ]

                    if i < len(lignes_filtrees) - 1:
                        for _ in range(2):
                            df_sortie.loc[len(df_sortie)] = [""] * len(entete)

                df_sortie = pd.concat([pd.DataFrame([entete], columns=entete), df_sortie], ignore_index=True)
                nom_fichier = os.path.join(dossier_sortie, f"AF_{identifiant.strip()}.xlsx")

                with pd.ExcelWriter(nom_fichier, engine='openpyxl') as writer:
                    df_sortie.to_excel(writer, index=False, header=False, startrow=5, startcol=0)
                    worksheet = writer.sheets['Sheet1']
                    worksheet['A1'] = "Rotation avion"
                    worksheet['A4'] = f"Matricule Avion: {identifiant}"
                    worksheet['B4'] = f"Type d'avion: {', '.join(map(str, valeurs_colonne_C))}"

            except Exception as e:
                print(f"Erreur pour {identifiant} : {e}")

            # Mise à jour de la barre et du canvas
            progress_bar["value"] = idx + 1
            pourcentage = int((idx + 1) / total_avions * 100)
            label_progress.config(text=f"Progression : {pourcentage}%")

            # Déplacement de l'avion sur le canvas
            new_x = int((idx + 1) / total_avions * canvas_width)
            canvas.coords(avion_item, new_x, 10)
            root.update_idletasks()
            time.sleep(0.05)

        messagebox.showinfo("Succès", "Vos données sont prêtes Captain ! 🚀\n\nDossier : " + dossier_sortie)
        ouvrir_dossier(dossier_sortie)

    except Exception as e:
        messagebox.showerror("Erreur", str(e))

# === Interface graphique ===
def lancer_interface():
    root = tk.Tk()
    root.title("Rotation vols Air France - Générateur")
    root.geometry("700x400")

    file_path_var = tk.StringVar()
    dossier_var = tk.StringVar()

    def choisir_fichier():
        chemin = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if chemin:
            file_path_var.set(chemin)

    def choisir_dossier():
        dossier = filedialog.askdirectory()
        if dossier:
            dossier_var.set(dossier)

    def lancer_traitement_thread():
        if not file_path_var.get() or not dossier_var.get():
            messagebox.showerror("Erreur", "Veuillez sélectionner un fichier et un dossier.")
            return
        threading.Thread(
            target=traiter_fichier,
            args=(file_path_var.get(), dossier_var.get(), progress_bar, label_progress, canvas, avion_item, root),
            daemon=True
        ).start()

    tk.Label(root, text="Fichier Excel source :").pack(pady=5)
    tk.Entry(root, textvariable=file_path_var, width=60).pack()
    tk.Button(root, text="Choisir fichier", command=choisir_fichier).pack(pady=5)

    tk.Label(root, text="Dossier de sortie :").pack(pady=5)
    tk.Entry(root, textvariable=dossier_var, width=60).pack()
    tk.Button(root, text="Choisir dossier", command=choisir_dossier).pack(pady=5)

    tk.Button(root, text="🚀 Lancer le traitement 🚀", command=lancer_traitement_thread, bg="green", fg="white").pack(pady=15)

    # Canvas pour animation avion
    canvas = tk.Canvas(root, width=600, height=30, bg="white")
    canvas.pack(pady=10)
    avion_item = canvas.create_text(0, 10, text="✈️", font=("Arial", 16), anchor="w")

    progress_bar = ttk.Progressbar(root, orient="horizontal", length=600, mode="determinate")
    progress_bar.pack(pady=5)

    label_progress = tk.Label(root, text="Progression : 0%")
    label_progress.pack()

    root.mainloop()

if __name__ == "__main__":
    lancer_interface()