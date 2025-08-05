import subprocess
import platform
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd
from docx import Document
from datetime import datetime
import sys

CONFIG_PATH = "config.txt"


def lire_chemin_modele():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            chemin = f.read().strip()
            if os.path.exists(chemin):
                return chemin
    return ""


def sauvegarder_chemin_modele(chemin):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        f.write(chemin)


try:
    import openpyxl
except ImportError:
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror(
        "Dépendance manquante",
        "Le module 'openpyxl' est requis pour lire les fichiers Excel.\n\n"
        "Veuillez l’installer avec la commande suivante dans votre terminal :\n\n"
        "pip install openpyxl"
    )
    sys.exit(1)


def remplir_modele(modele_path, donnees):
    doc = Document(modele_path)
    for p in doc.paragraphs:
        for key, value in donnees.items():
            balise = f"{{{{{key}}}}}"
            if balise in p.text:
                p.text = p.text.replace(balise, str(value))
    return doc


def traiter_fichier_excel(fichier_excel, fichier_modele, dossier_genere_var, bouton_afficher, label_progression, progress_bar, fenetre):
    if not os.path.exists(fichier_modele):
        messagebox.showerror(
            "Erreur", f"Le fichier modèle est introuvable :\n{fichier_modele}")
        return

    if not fichier_excel.lower().endswith(".xlsx"):
        messagebox.showerror(
            "Erreur", "Veuillez sélectionner un fichier Excel valide (.xlsx)")
        return

    try:
        df = pd.read_excel(fichier_excel)
    except Exception as e:
        messagebox.showerror("Erreur lors de la lecture Excel", str(e))
        return

    if df.empty:
        messagebox.showinfo("Info", "Le fichier Excel est vide.")
        return

    now = datetime.now()
    dossier_sortie = os.path.join(os.getcwd(), f"{now.year}-{now.month:02d}")
    os.makedirs(dossier_sortie, exist_ok=True)

    total = len(df)
    for index, row in df.iterrows():
        donnees = {
            "société": row.get("Société", ""),
            "contact": row.get("Contact", ""),
            "mail": row.get("Mail", ""),
            "portable": str(row.get("Portable", "")),
            "adresse": row.get("Adresse", ""),
            "cp": str(row.get("Cp", "")),
            "ville": row.get("Ville", "")
        }

        try:
            doc = remplir_modele(fichier_modele, donnees)
            nom_fichier = f"{donnees['société']}_{donnees['contact'].split()[0]}.docx"
            chemin_sortie = os.path.join(dossier_sortie, nom_fichier)
            doc.save(chemin_sortie)
        except Exception as e:
            messagebox.showwarning("Erreur de traitement",
                                   f"Erreur sur la ligne {index + 1} : {e}")

        label_progression.config(text=f"Traitement : {index + 1} / {total}")
        fenetre.update_idletasks()
        progress_bar["value"] = int((index + 1) / total * 100)

    dossier_genere_var.set(dossier_sortie)
    bouton_afficher.pack()
    label_progression.config(text="✅ Traitement terminé !")
    messagebox.showinfo(
        "Terminé", f"{total} document(s) généré(s) dans :\n{dossier_sortie}")


def ouvrir_dossier(path):
    systeme = platform.system()
    if systeme == "Windows":
        os.startfile(path)
    elif systeme == "Darwin":  # macOS
        subprocess.run(["open", path])
    elif systeme == "Linux":
        subprocess.run(["xdg-open", path])


def lancer_gui():
    fenetre = tk.Tk()
    fenetre.title("Générateur Word depuis Excel")
    fenetre.geometry("650x480")
    fenetre.configure(bg="#f8f8f8")

    style = ttk.Style(fenetre)
    style.theme_use("clam")
    style.configure("TNotebook", background="#f8f8f8", borderwidth=0)
    style.configure("TNotebook.Tab", padding=[12, 8], font=("Segoe UI", 11))
    style.configure("TLabel", background="#f8f8f8", font=("Segoe UI", 10))
    style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)
    style.configure("TProgressbar", thickness=20)

    notebook = ttk.Notebook(fenetre)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    # --- Onglet Générateur ---
    frame_gen = ttk.Frame(notebook)
    notebook.add(frame_gen, text="Générateur")

    chemin_excel = tk.StringVar()
    chemin_modele = tk.StringVar()
    dossier_genere = tk.StringVar()

    def choisir_excel():
        fichier = filedialog.askopenfilename(
            title="Choisir un fichier Excel",
            filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
        )
        if fichier:
            chemin_excel.set(fichier)

    def choisir_modele_word():
        fichier = filedialog.askopenfilename(
            title="Choisir un fichier Word modèle",
            filetypes=[("Fichiers Word", "*.docx")]
        )
        if fichier:
            chemin_modele.set(fichier)
            sauvegarder_chemin_modele(fichier)
            label_info.config(
                text=f"Modèle Word par défaut chargé : {fichier}", foreground="green")

    def lancer_traitement():
        if not chemin_excel.get():
            messagebox.showwarning(
                "Attention", "Veuillez importer un fichier Excel.")
            return
        if not chemin_modele.get():
            messagebox.showwarning(
                "Attention", "Veuillez importer un modèle Word.")
            return
        bouton_ouvrir_dossier.pack_forget()
        traiter_fichier_excel(chemin_excel.get(), chemin_modele.get(),
                              dossier_genere, bouton_ouvrir_dossier, label_progression, progress_bar, fenetre)

    # Widgets onglet Générateur
    ttk.Label(frame_gen, text="Fichier Excel sélectionné :").pack(anchor="w", pady=5)
    ttk.Entry(frame_gen, textvariable=chemin_excel,
              width=80, state='readonly').pack(pady=2)

    ttk.Button(frame_gen, text="📄 Importer un fichier Excel",
               command=choisir_excel).pack(pady=5)

    ttk.Label(frame_gen, text="Fichier Word modèle sélectionné :").pack(anchor="w", pady=5)
    ttk.Entry(frame_gen, textvariable=chemin_modele,
              width=80, state='readonly').pack(pady=2)

    ttk.Button(frame_gen, text="📄 Importer un modèle Word",
               command=choisir_modele_word).pack(pady=5)

    ttk.Button(frame_gen, text="🚀 Lancer le traitement",
               command=lancer_traitement, style="TButton").pack(pady=12)

    progress_bar = ttk.Progressbar(
        frame_gen, orient='horizontal', length=450, mode='determinate')
    progress_bar.pack(pady=5)

    label_progression = ttk.Label(frame_gen, text="", foreground="blue")
    label_progression.pack(pady=3)

    bouton_ouvrir_dossier = ttk.Button(
        frame_gen,
        text="📂 Ouvrir le dossier généré",
        command=lambda: ouvrir_dossier(dossier_genere.get())
    )
    bouton_ouvrir_dossier.pack(pady=10)
    bouton_ouvrir_dossier.pack_forget()

    label_info = ttk.Label(frame_gen, text="", foreground="green", wraplength=600, justify="left")
    label_info.pack(pady=5)

    # Charger chemin modèle sauvegardé
    chemin_modele_sauvegarde = lire_chemin_modele()
    if chemin_modele_sauvegarde:
        chemin_modele.set(chemin_modele_sauvegarde)
        label_info.config(
            text=f"Modèle Word par défaut chargé : {chemin_modele_sauvegarde}")

    # --- Onglet À propos ---
    frame_about = ttk.Frame(notebook)
    notebook.add(frame_about, text="À propos")

    ttk.Label(frame_about, text="Développé par :", font=("Segoe UI", 12, "bold")).pack(pady=(30,5))
    ttk.Label(frame_about, text="Edwin Dijeont", font=("Segoe UI", 11)).pack(pady=2)
    ttk.Label(frame_about, text="Email :", font=("Segoe UI", 12, "bold")).pack(pady=(20,5))
    ttk.Label(frame_about, text="edwin.d899@gmail.com", font=("Segoe UI", 11), foreground="blue").pack(pady=2)

    fenetre.mainloop()


if __name__ == "__main__":
    lancer_gui()
