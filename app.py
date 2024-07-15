import customtkinter as ctk
import tkinter.filedialog
import shutil
import os
import re
import fitz  # PyMuPDF pour manipuler les PDF
import subprocess  # Pour ouvrir l'explorateur de fichiers
import threading  # Pour gérer la durée d'affichage du message

# Définir le thème et l'apparence de l'interface
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Chemin vers le dossier Facture
facture_folder = os.path.join(os.path.expanduser("~"), "Desktop", "DossierOmnia", "Facture")
# Chemin vers le dossier BonCommande
bon_commande_folder = os.path.join(os.path.expanduser("~"), "Desktop", "DossierOmnia", "BonCommande")

# Créer les dossiers s'ils n'existent pas
os.makedirs(facture_folder, exist_ok=True)
os.makedirs(bon_commande_folder, exist_ok=True)

# Créer l'interface principale
interface = ctk.CTk()
interface.geometry("800x450")

# Fonction pour remplacer le texte dans un fichier PDF
def replace_text_in_pdf(input_pdf_path, output_pdf_path, old_text, new_text_prefix):
    # Ouvrir le fichier PDF
    pdf_document = fitz.open(input_pdf_path)

    for page_num in range(len(pdf_document)):
        # Charger chaque page du document
        page = pdf_document.load_page(page_num)
        # Chercher toutes les instances de l'ancien texte sur la page
        text_instances = page.search_for(old_text)

        for inst in text_instances:
            # Dessiner un rectangle blanc sur l'ancien texte pour l'effacer
            page.draw_rect(inst, color=(1, 1, 1), fill=(1, 1, 1))

            # Extraire le code après "Facture FA"
            old_text_with_code = page.get_textbox(inst)
            match = re.search(r'Facture FA(\d+-\d+)', old_text_with_code, re.IGNORECASE)
            if match:
                code = match.group(1)  # Récupérer le code (par exemple "2406-7437")
            else:
                code = ""  # Par défaut, si aucun code n'est trouvé

            # Créer le nouveau texte avec "Bon de commande BC" suivi du code
            new_text = f"{new_text_prefix} {code}"  # Ajouter un espace avant le code

            # Ajouter le texte en gras
            x, y, _, _ = inst
            new_x = x - 62  # Ajuster cette valeur pour déplacer plus ou moins
            new_y = y + 14
            page.insert_text((new_x, new_y), new_text, fontsize=12, fontname="Helvetica-Bold", color=(0, 0, 0))

    # Sauvegarder le fichier PDF modifié
    pdf_document.save(output_pdf_path)
    pdf_document.close()

# Fonction pour charger et enregistrer la facture
def upload_facture():
    # Sélectionner le fichier PDF
    file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        # Nom du fichier sans le chemin complet
        file_name = os.path.basename(file_path)

        # Chemin pour la facture (dossier Facture)
        facture_path = os.path.join(facture_folder, file_name)
        # Chemin pour le bon de commande (dossier BonCommande)
        bon_commande_path = os.path.join(bon_commande_folder, f"Bon_Commande_{file_name}")

        # Copier la facture dans le dossier Facture
        shutil.copy(file_path, facture_path)
        
        # Remplacer le texte dans le bon de commande et enregistrer dans le dossier BonCommande
        replace_text_in_pdf(file_path, bon_commande_path, "Facture FA", "Bon de commande BC")
        
        # Afficher un message de succès temporaire dans le cadre en bas
        show_success_message("Bien,la transformation de la facture vers une bon de commande réussie!, vous pouvez cliquer sur ouvrir l'emplacement ")

        # Activer le bouton pour ouvrir l'emplacement
        button_open_folder.configure(state=ctk.NORMAL)

# Fonction pour afficher un message de succès temporaire dans le cadre en bas
def show_success_message(message, duration=3):
    # Afficher le message dans une étiquette en bas du cadre
    label_success.configure(text=message)

    # Désactiver le bouton pour éviter l'affichage simultané de plusieurs messages
    button_upload.configure(state=ctk.DISABLED)

    # Décompte pour réactiver le bouton après 'duration' secondes
    def restore_button_state():
        button_upload.configure(state=ctk.NORMAL)  # Réactiver le bouton
        label_success.configure(text="")  # Effacer le message de succès

    # Démarrer un thread pour gérer la durée d'affichage du message
    threading.Timer(duration, restore_button_state).start()

# Fonction pour ouvrir l'emplacement des fichiers
def open_folder():
    # Ouvrir l'explorateur de fichiers à l'emplacement du dossier BonCommande
    subprocess.Popen(f'explorer "{bon_commande_folder}"')

# Créer le cadre principal
frame = ctk.CTkFrame(master=interface)
frame.pack(pady=20, padx=20, fill='both', expand=True)

# Ajouter une étiquette de bienvenue
label_welcome = ctk.CTkLabel(master=frame, text="Bienvenue")
label_welcome.configure(text_color='white', font=('Arial', 25, 'bold'))
label_welcome.pack(pady=20)

# Ajouter un bouton pour charger la facture
button_upload = ctk.CTkButton(master=frame, text="Cliquer ici pour choisir une Facture et le transformer en Bon de commande", command=upload_facture)
button_upload.configure(text_color='white', font=('Arial', 17))
button_upload.pack(pady=20)

# Ajouter un bouton pour ouvrir l'emplacement
button_open_folder = ctk.CTkButton(master=frame, text="Ouvrir l'emplacement", command=open_folder)
button_open_folder.configure(text_color='white', font=('Arial', 12), state=ctk.DISABLED)
button_open_folder.pack(pady=10)

# Ajouter une étiquette pour afficher les messages de succès en bas
label_success = ctk.CTkLabel(master=interface, text="", text_color='green', font=('Arial', 12))
label_success.pack(side='bottom', pady=20)

# Lancer l'interface
interface.mainloop()
