import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk

def redimensionner_image(image_path, nouvelle_taille):
    # Fonction pour redimensionner une image
    img = Image.open(image_path)
    img_redimensionnee = img.resize(nouvelle_taille)
    return ImageTk.PhotoImage(img_redimensionnee)

def renommer_fichiers():
    # Fonction pour renommer les fichiers dans un dossier en utilisant un fichier Excel

    # Obtenez le chemin du dossier d'images
    dossier_images = filedialog.askdirectory(title="Sélectionnez le dossier de fichiers à renommer : ")

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

        # Vérifiez si la colonne 'refFournisseur' existe dans le DataFrame
        if 'refFournisseur' in df_excel.columns:
            # Convertissez la colonne 'refFournisseur' en chaînes de caractères
            df_excel['refFournisseur'] = df_excel['refFournisseur'].astype(str)

            # Barre de progression
            progress_bar = ttk.Progressbar(fenetre, orient="horizontal", length=300, mode="determinate")
            progress_bar.pack(pady=10)

            # Widget Text pour les logs de la console
            logs_text = tk.Text(fenetre, height=10, width=60)
            logs_text.pack(pady=10)
            logs_text.insert(tk.END, "Etats du renommage :\n\n")

            # Parcourez tous les fichiers dans le dossier d'images
            total_files = len([f for f in os.listdir(dossier_images) if os.path.isfile(os.path.join(dossier_images, f))])
            progress_bar["maximum"] = total_files

            for i, fichier_image in enumerate(os.listdir(dossier_images)):
                chemin_image = os.path.join(dossier_images, fichier_image)

                # Mise à jour de la barre de progression
                progress_bar["value"] = i + 1
                fenetre.update_idletasks()

                # Obtenez le nom de fichier sans extension et l'extension
                nom_fichier_sans_extension, extension = os.path.splitext(fichier_image)


                # Vérifiez si le fichier contient un tiret
                if '-' in nom_fichier_sans_extension:
                    # Divisez le nom de fichier en deux parties (avant et après le tiret)
                    avant_tiret, apres_tiret = nom_fichier_sans_extension.split('-', 1)

                    # Recherchez le nom dans la colonne 'refFournisseur' du fichier Excel
                    correspondance = df_excel[df_excel['refFournisseur'] == avant_tiret]

                    # Si une correspondance est trouvée, renommez le fichier avec la partie après le tiret et ajoutez les tirets
                    if not correspondance.empty:
                        nouvelle_valeur = correspondance.iloc[0]['refArticle']
                        nouveau_nom = f"{nouvelle_valeur}-{apres_tiret}{extension}"  # Utilisez l'extension d'origine
                        nouveau_chemin = os.path.join(dossier_images, nouveau_nom)
                        
                        # Gestion du conflit de nom de fichier
                        counter = 1
                        while os.path.exists(nouveau_chemin):
                            nouveau_nom = f"{nouvelle_valeur}-{apres_tiret}_{counter}{extension}"
                            nouveau_chemin = os.path.join(dossier_images, nouveau_nom)
                            counter += 1
                        # Renommez le fichier
                        os.rename(chemin_image, nouveau_chemin)

                        log_message = f"Le fichier {fichier_image} a été renommé en {nouveau_nom}\n"
                        logs_text.insert(tk.END, log_message)
                        print(log_message)
                    else:
                        log_message = f'Aucune correspondance pour {avant_tiret} !\n'
                        logs_text.insert(tk.END, log_message)
                        print(log_message)
                else:
                    # Si le fichier ne contient pas de tiret, renommez-le en utilisant la valeur dans la colonne 'refFournisseur' et ajoutez les tirets
                    correspondance = df_excel[df_excel['refFournisseur'] == nom_fichier_sans_extension]
                    if not correspondance.empty:
                        nouvelle_valeur = correspondance.iloc[0]['refArticle']
                        nouveau_nom = f"{nouvelle_valeur}{extension}"  # Utilisez l'extension d'origine
                        nouveau_chemin = os.path.join(dossier_images, nouveau_nom)

                        # Gestion du conflit de nom de fichier
                        counter = 1
                        while os.path.exists(nouveau_chemin):
                            nouveau_nom = f"{nouvelle_valeur}_{counter}{extension}"
                            nouveau_chemin = os.path.join(dossier_images, nouveau_nom)
                            counter += 1
                        # Renommez le fichier
                        os.rename(chemin_image, nouveau_chemin)

                        log_message = f"Le fichier {fichier_image} a été renommé en {nouveau_nom}\n"
                        logs_text.insert(tk.END, log_message)
                        print(log_message)
                    else:
                        log_message = f'Aucune correspondance pour {nom_fichier_sans_extension} !\n'
                        logs_text.insert(tk.END, log_message)
                        print(log_message)

            messagebox.showinfo("Opération terminée", "Le renommage des fichiers est terminé avec succès!")

        else:
            messagebox.showwarning("Avertissement", "La colonne 'refFournisseur' n'a pas été trouvée dans le fichier Excel.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur s'est produite : {str(e)}")

    # Réinitialisez la barre de progression
    progress_bar.destroy()
    # Ajoutez un bouton OK pour fermer la fenêtre après l'opération
    ok_button = tk.Button(fenetre, text="OK", command=fenetre.destroy)
    ok_button.pack(pady=20)

# Créez une fenêtre Tkinter
fenetre = tk.Tk()
fenetre.title("Renommer des fichiers")
fenetre.geometry("600x400")

# Cachez la barre de titre et les bordures de la fenêtre
fenetre.attributes('-fullscreen', True)

# Utilisez un thème moderne avec ttk
style = ttk.Style()
style.theme_use('clam')  # 'clam' est un thème moderne, vous pouvez choisir d'autres thèmes selon vos préférences

# Ajoutez une icône personnalisée à la fenêtre
icone_path = os.path.join(os.path.dirname(__file__), "logo_bernabee.ico")
fenetre.wm_iconbitmap(icone_path, default=icone_path)

# Redimensionnez l'image avant de l'afficher
image_path = os.path.join(os.path.dirname(__file__), "logo_bernabee.png")
nouvelle_taille = (250, 250)  # Spécifiez la nouvelle taille souhaitée
image_redimensionnee = redimensionner_image(image_path, nouvelle_taille)
label_image = ttk.Label(fenetre, image=image_redimensionnee)
label_image.pack(pady=10)

# Créez un bouton pour lancer l'opération
bouton_lancer = ttk.Button(fenetre, text="Lancer l'opération", width=20, command=renommer_fichiers)
bouton_lancer.pack(pady=20)

# Ajoutez des boutons pour fermer et réduire la fenêtre
fermer_button = tk.Button(fenetre, bg='#FF6961', width=20, height=2, text="Fermer", command=fenetre.destroy)
fermer_button.pack(pady=20)

def reduire_fenetre():
    fenetre.iconify()

reduire_button = tk.Button(fenetre, bg='#CEC', width=20, height=2, text="Réduire", command=reduire_fenetre)
reduire_button.pack(pady=20)

# Démarrez la boucle principale Tkinter
fenetre.mainloop()
