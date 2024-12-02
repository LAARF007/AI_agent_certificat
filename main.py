import os
from docx import Document
from datetime import datetime
import pypandoc
import io

# Dossier de sortie pour les fichiers générés
output_folder = "output"
os.makedirs(output_folder, exist_ok=True)

def remplir_template_docx(nom, prenom, code_massar, lieu):
    # Charger le modèle Word
    template_path = "template.docx"  # Le modèle Word
    output_path = os.path.join(output_folder, f"attestation_{nom}_{prenom}.docx")  # Nom du fichier généré
    doc = Document(template_path)

    # Remplacer les champs dans le modèle
    for paragraph in doc.paragraphs:
        if "{{NOM}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{NOM}}", nom)
        if "{{PRENOM}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{PRENOM}}", prenom)
        if "{{CODE_MASSAR}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{CODE_MASSAR}}", code_massar)
        if "{{LIEU}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{LIEU}}", lieu)
        if "{{DATE}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{DATE}}", datetime.now().strftime("%d/%m/%Y"))

    # Sauvegarder le fichier Word généré dans le dossier output
    doc.save(output_path)
    print(f"Attestation générée : {output_path}")
    return output_path

# def convertir_docx_en_pdf(docx_path):
#     # Utilisation de pypandoc pour convertir en PDF
#     pdf_path = docx_path.replace(".docx", ".pdf")
#     try:
#         output = pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)
#         print(f"Le fichier PDF a été généré : {pdf_path}")
#         return pdf_path
#     except Exception as e:
#         print(f"Erreur lors de la conversion en PDF : {e}")
#         return None
def convertir_docx_en_pdf(docx_path):
    # Générer le nom du fichier PDF à partir du fichier DOCX
    pdf_path = docx_path.replace(".docx", ".pdf")

    try:
        # Utilisation de pypandoc pour convertir en PDF
        output = pypandoc.convert_file(docx_path, 'pdf', outputfile=pdf_path)

        # Lire le fichier PDF et le retourner sous forme de bytes
        with open(pdf_path, "rb") as pdf_file:
            pdf_bytes = io.BytesIO(pdf_file.read())

        print(f"Le fichier PDF a été généré : {pdf_path}")
        return pdf_bytes

    except Exception as e:
        print(f"Erreur lors de la conversion en PDF : {e}")
        return None

def main():
    # Saisie des informations
    nom = input("Entrez le nom : ")
    prenom = input("Entrez le prénom : ")
    code_massar = input("Entrez le code Massar : ")
    lieu = input("Entrez le lieu : ")

    # Générer l'attestation Word
    fichier_docx = remplir_template_docx(nom, prenom, code_massar, lieu)

    # Convertir le fichier Word en PDF
    fichier_pdf = convertir_docx_en_pdf(fichier_docx)

    if fichier_pdf:
        print(f"Le fichier PDF a été enregistré dans : {fichier_pdf}")
    else:
        print("La conversion en PDF a échoué.")

if __name__ == "__main__":
    main()
