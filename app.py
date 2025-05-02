import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# Variables globales
df_remplacements = pd.DataFrame(columns=["Ancien", "Nouveau"])

# Fonction pour parcourir un fichier Excel
def parcourir_fichier():
    fichier = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx")])
    if fichier:
        entree_fichier.delete(0, tk.END)
        entree_fichier.insert(0, fichier)
        charger_tableau_remplacements(fichier)

# Chargement depuis Excel
def charger_tableau_remplacements(fichier_excel):
    global df_remplacements
    try:
        df_remplacements = pd.read_excel(fichier_excel)
        if "Ancien" not in df_remplacements.columns or "Nouveau" not in df_remplacements.columns:
            messagebox.showerror("Erreur", "Le fichier Excel doit contenir les colonnes 'Ancien' et 'Nouveau'.")
            return
        rafraichir_tableau()
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de la lecture du fichier Excel :\n{e}")

# Mise à jour de l'affichage
def rafraichir_tableau():
    for row in tableau.get_children():
        tableau.delete(row)
    for _, row in df_remplacements.iterrows():
        tableau.insert("", tk.END, values=(row["Ancien"], row["Nouveau"]))

# Ajouter une ligne
def ajouter_ligne():
    global df_remplacements
    ancien = entree_ancien.get().strip()
    nouveau = entree_nouveau.get().strip()
    if ancien and nouveau:
        df_remplacements = pd.concat([df_remplacements, pd.DataFrame([[ancien, nouveau]], columns=["Ancien", "Nouveau"])], ignore_index=True)
        rafraichir_tableau()
        entree_ancien.delete(0, tk.END)
        entree_nouveau.delete(0, tk.END)

# Supprimer une ligne
def supprimer_ligne():
    global df_remplacements
    selected_item = tableau.selection()
    if selected_item:
        item = tableau.item(selected_item)
        valeurs = item["values"]
        df_remplacements = df_remplacements[~((df_remplacements["Ancien"] == valeurs[0]) & (df_remplacements["Nouveau"] == valeurs[1]))]
        rafraichir_tableau()

# Modifier une ligne
def modifier_ligne():
    global df_remplacements
    selected_item = tableau.selection()
    if selected_item:
        item = tableau.item(selected_item)
        ancien_val, nouveau_val = item["values"]
        nouveau_ancien = entree_ancien.get().strip()
        nouveau_nouveau = entree_nouveau.get().strip()
        if nouveau_ancien and nouveau_nouveau:
            idx = df_remplacements[(df_remplacements["Ancien"] == ancien_val) & (df_remplacements["Nouveau"] == nouveau_val)].index
            if not idx.empty:
                df_remplacements.loc[idx[0], "Ancien"] = nouveau_ancien
                df_remplacements.loc[idx[0], "Nouveau"] = nouveau_nouveau
                rafraichir_tableau()
                entree_ancien.delete(0, tk.END)
                entree_nouveau.delete(0, tk.END)

# Lancer le remplacement dans un XML
def lancer_remplacement():
    if df_remplacements.empty:
        messagebox.showerror("Erreur", "Aucun remplacement à effectuer.")
        return

    fichier_xml = filedialog.askopenfilename(filetypes=[("Fichiers XML", "*.xml")])
    if not fichier_xml:
        return

    try:
        with open(fichier_xml, 'r', encoding='utf-8') as f:
            contenu = f.read()

        log = ""
        for index, row in df_remplacements.iterrows():
            ancien = str(row["Ancien"]).strip()
            nouveau = str(row["Nouveau"]).strip()

            if ancien in contenu:
                if ancien == nouveau:
                    log += f"🔁 Identique : {ancien}\n"
                else:
                    contenu = contenu.replace(ancien, nouveau)
                    log += f"✅ Remplacé : {ancien} → {nouveau}\n"

        # Choisir le nom et l’emplacement du fichier de sortie
        nouveau_fichier = filedialog.asksaveasfilename(
            defaultextension=".xml",
            filetypes=[("Fichiers XML", "*.xml")],
            title="Enregistrer sous...",
            initialfile=os.path.basename(fichier_xml).replace(".xml", "_modifié.xml")
        )
        if not nouveau_fichier:
            return

        with open(nouveau_fichier, 'w', encoding='utf-8') as f:
            f.write(contenu)

        zone_resultats.delete("1.0", tk.END)
        zone_resultats.insert(tk.END, log)
        messagebox.showinfo("Succès", f"Remplacement terminé !\nFichier sauvegardé : {nouveau_fichier}")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors du remplacement :\n{e}")

# Interface graphique
fenetre = tk.Tk()
fenetre.title("🛠️ Remplacement Intelligent XML SCF")

# Sélection fichier Excel
frame_haut = tk.Frame(fenetre)
frame_haut.pack(pady=10)

tk.Label(frame_haut, text="Fichier Excel de remplacement :").pack(side=tk.LEFT)
entree_fichier = tk.Entry(frame_haut, width=60)
entree_fichier.pack(side=tk.LEFT, padx=5)
tk.Button(frame_haut, text="📂 Parcourir", command=parcourir_fichier).pack(side=tk.LEFT)

# Lancer remplacement
tk.Button(fenetre, text="⚙️ Lancer le remplacement dans XML", command=lancer_remplacement, bg="lightgreen").pack(pady=10)

# Tableau interactif
tk.Label(fenetre, text="📝 Tableau des remplacements (Excel)").pack()
colonnes = ("Ancien", "Nouveau")
tableau = ttk.Treeview(fenetre, columns=colonnes, show="headings", height=10)
for col in colonnes:
    tableau.heading(col, text=col)
    tableau.column(col, width=400)
tableau.pack()

# Zone de modification
frame_modif = tk.Frame(fenetre)
frame_modif.pack(pady=5)

tk.Label(frame_modif, text="Ancien:").grid(row=0, column=0)
entree_ancien = tk.Entry(frame_modif, width=40)
entree_ancien.grid(row=0, column=1)

tk.Label(frame_modif, text="Nouveau:").grid(row=0, column=2)
entree_nouveau = tk.Entry(frame_modif, width=40)
entree_nouveau.grid(row=0, column=3)

tk.Button(frame_modif, text="➕ Ajouter", command=ajouter_ligne, bg="#d0f0c0").grid(row=1, column=1, pady=5)
tk.Button(frame_modif, text="✏️ Modifier", command=modifier_ligne, bg="#fef9c3").grid(row=1, column=2)
tk.Button(frame_modif, text="❌ Supprimer", command=supprimer_ligne, bg="#fecaca").grid(row=1, column=3)

# Journal des remplacements
tk.Label(fenetre, text="🧾 Journal des remplacements").pack()
zone_resultats = tk.Text(fenetre, height=15, width=100, bg="#e8f5e9")
zone_resultats.pack()

fenetre.mainloop()
