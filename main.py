import os
from pptx import Presentation
import comtypes.client
from datetime import datetime
import smtplib
from email.message import EmailMessage


# Chemins
base_path = os.path.dirname(os.path.abspath(__file__))
template_pptx = os.path.join(base_path, "attesstaion_temp.docx")
output_folder = os.path.join(base_path, "output")
os.makedirs(output_folder, exist_ok=True)


def ppt_to_pdf(input_pptx, output_pdf):
    """Convertir un fichier PPTX en PDF."""
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    if output_pdf[-3:] != "pdf":
        output_pdf = output_pdf + ".pdf"
    deck = powerpoint.Presentations.Open(input_pptx, WithWindow=False)
    deck.ExportAsFixedFormat(output_pdf, FixedFormatType=2)
    deck.Close()
    powerpoint.Quit()


def generate_attestation(data):
    """Générer une attestation personnalisée."""
    prs = Presentation(template_pptx)

    # Remplacement des placeholders
    placeholders = {
        "Placeholder_Name": data["name"],
        "Placeholder_CodeMassar": data["code_massar"],
        "Placeholder_Ville": data["ville"],
        "Placeholder_Date": data["date"],
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for placeholder, value in placeholders.items():
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, value)

    # Enregistrer le fichier PPTX et le convertir en PDF
    output_pptx = os.path.join(output_folder, f"{data['code_massar']}.pptx")
    prs.save(output_pptx)
    output_pdf = os.path.join(output_folder, f"{data['code_massar']}.pdf")
    ppt_to_pdf(output_pptx, output_pdf)

    print(f"Attestation générée avec succès : {output_pdf}")
    return output_pdf  # Retourne le chemin du PDF généré


def send_email_with_attachment(email, pdf_path, subject="Attestation de Scolarité"):
    """Envoyer un e-mail avec le PDF en pièce jointe."""
    # Configuration SMTP
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    sender_email = "laarfadel33@gmail.com"  # Remplacez par votre e-mail
    sender_password = "votre_mot_de_passe"  # Remplacez par votre mot de passe

    # Composer l'e-mail
    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = email
    message["Subject"] = subject
    message.set_content("Veuillez trouver ci-joint votre attestation de scolarité.")

    # Ajouter le PDF en pièce jointe
    with open(pdf_path, "rb") as file:
        file_data = file.read()
        file_name = os.path.basename(pdf_path)
        message.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)

    # Envoyer l'e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()  # Démarrer la connexion sécurisée
        server.login(sender_email, sender_password)
        server.send_message(message)

    print(f"E-mail envoyé avec succès à {email}.")


if __name__ == "__main__":
    print("Veuillez entrer les informations pour l'attestation.")

    # Collecter les données via input()
    user_data = {
        "name": input("Nom complet : "),
        "code_massar": input("Code Massar : "),
        "ville": input("Ville : "),
        "date": input("Date (format JJ/MM/AAAA) : ") or datetime.now().strftime("%d/%m/%Y"),
    }

    user_email = input("Adresse e-mail : ")

    # Générer l'attestation
    pdf_path = generate_attestation(user_data)

    # Envoyer l'attestation par e-mail
    send_email_with_attachment(user_email, pdf_path)
