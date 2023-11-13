import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def renommer_fichiers():
    # Fonction pour renommer les fichiers

    # Obtenez le chemin du dossier d'images
    dossier_images = filedialog.askdirectory(title="Sélectionnez le dossier d'images")

    if not dossier_images:
        # L'utilisateur a annulé la sélection du dossier
        return

    # Obtenez le chemin du fichier Excel
    fichier_excel = filedialog.askopenfilename(title="Sélectionnez le fichier Excel", filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])

    if not fichier_excel:
        # L'utilisateur a annulé la sélection du fichier Excel
        return

    try:
        # Chargez le fichier Excel dans une dataframe pandas
        df_excel = pd.read_excel(fichier_excel)

        # Vérifiez si la colonne 'seq' existe dans le DataFrame
        if 'seq' in df_excel.columns:
            # Convertissez la colonne 'seq' en chaînes de caractères
            df_excel['seq'] = df_excel['seq'].astype(str)

            # Parcourez tous les fichiers dans le dossier d'images
            for fichier_image in os.listdir(dossier_images):
                chemin_image = os.path.join(dossier_images, fichier_image)

                # Vérifiez si le fichier est un fichier image
                if os.path.isfile(chemin_image) and fichier_image.lower().endswith(('.png', '.jpg', '.jpeg')):
                    # Obtenez le nom de fichier sans extension et l'extension
                    nom_fichier_sans_extension, extension = os.path.splitext(fichier_image)

                    # Vérifiez si le fichier contient un tiret
                    if '-' in nom_fichier_sans_extension:
                        # Divisez le nom de fichier en deux parties (avant et après le tiret)
                        avant_tiret, apres_tiret = nom_fichier_sans_extension.split('-', 1)

                        # Recherchez le nom dans la colonne 'seq' du fichier Excel
                        correspondance = df_excel[df_excel['ref'] == avant_tiret]

                        # Si une correspondance est trouvée, renommez le fichier avec la partie après le tiret et ajoutez les tirets
                        if not correspondance.empty:
                            nouvelle_valeur = correspondance.iloc[0]['ref']
                            nouveau_nom = f"{nouvelle_valeur}-{apres_tiret}{extension}"  # Utilisez l'extension d'origine
                            nouveau_chemin = os.path.join(dossier_images, nouveau_nom)

                            # Renommez le fichier
                            os.rename(chemin_image, nouveau_chemin)

                            print(f"Le fichier {fichier_image} a été renommé en {nouveau_nom}")
                        else:
                            print(f'Aucune correspondance pour {avant_tiret} !')
                    else:
                        # Si le fichier ne contient pas de tiret, renommez-le en utilisant la valeur dans la colonne 'seq' et ajoutez les tirets
                        correspondance = df_excel[df_excel['seq'] == nom_fichier_sans_extension]
                        if not correspondance.empty:
                            nouvelle_valeur = correspondance.iloc[0]['ref']
                            nouveau_nom = f"{nouvelle_valeur}{extension}"  # Utilisez l'extension d'origine
                            nouveau_chemin = os.path.join(dossier_images, nouveau_nom)

                            # Renommez le fichier
                            os.rename(chemin_image, nouveau_chemin)

                            print(f"Le fichier {fichier_image} a été renommé en {nouveau_nom}")
                        else:
                            print(f'Aucune correspondance pour {nom_fichier_sans_extension} !')


            messagebox.showinfo("Opération terminée", "Le renommage des fichiers est terminé avec succès!")

        else:
            messagebox.showwarning("Avertissement", "La colonne 'seq' n'a pas été trouvée dans le fichier Excel.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

# Créez une fenêtre Tkinter
fenetre = tk.Tk()
fenetre.title("Renommer des fichiers")
fenetre.geometry("720x360")
# Créez un bouton pour lancer l'opération
bouton_lancer = tk.Button(fenetre, text="Lancer l'opération", command=renommer_fichiers)
bouton_lancer.pack(pady=20)

# Démarrez la boucle principale Tkinter
fenetre.mainloop()
