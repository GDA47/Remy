import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import zipfile
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from textwrap import wrap
import os

# --- Load static assets ---
LOGO_PATH = "logo1.PNG"
SIGNATURE_PATH = "signaturer.PNG"

def charger_image(path):
    if os.path.exists(path):
        return ImageReader(path)
    return None

logo_image = charger_image(LOGO_PATH)
signature_image = charger_image(SIGNATURE_PATH)

# --- Date formatting ---
def formater_date_lettres(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y")
    mois = ["janvier", "février", "mars", "avril", "mai", "juin",
            "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    return f"{date_obj.day} {mois[date_obj.month - 1]} {date_obj.year}"

def generer_pdf(nom, date_str, commune, code_postal, logo, signature):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largeur, hauteur = A4

    marge_gauche = 2 * cm
    marge_droite = largeur - 5 * cm

    # ✅ Logo lower so it's fully visible (moved down)
    if logo:
        c.drawImage(logo, x=marge_gauche, y=hauteur - 7 * cm,
                    width=largeur - 2 * cm, preserveAspectRatio=True, mask='auto')

    # ✅ Date slightly below logo
    c.setFont("Helvetica", 12)
    c.setFillColorRGB(0, 0, 0)  # black text
    c.drawString(marge_gauche, hauteur - 6.2 * cm, f"Agen, le {formater_date_lettres(date_str)}")

    # ✅ Title box below date
    box_top = hauteur - 7.5 * cm
    box_bottom = box_top - 2.4 * cm
    c.setStrokeColorRGB(0.3, 0.6, 0.3)  # green border
    c.setFillColorRGB(0.85, 0.95, 0.85)  # light green fill
    c.rect(marge_gauche, box_bottom, largeur - 4 * cm, 2.4 * cm, fill=1, stroke=1)

    # ✅ Title text (black, bold, two lines)
    c.setFont("Helvetica-Bold", 14)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(largeur / 2, box_top - 0.9 * cm,
        "Attestation de Suivi Technique")
    c.drawCentredString(largeur / 2, box_top - 1.7 * cm,
        "Pomme Production Fruitière Intégrée")

    # ✅ Body text using full width with equal margins
    c.setFont("Helvetica", 11)
    body = f"""J’atteste que {nom} à {commune.upper()} ({code_postal[:2]}) a souscrit à un suivi technique en Arboriculture auprès de notre chambre d’agriculture. 
A ce titre :
• Son verger est suivi au moins à 3 reprises durant l’année, avec une préconisation. 
• Il reçoit chaque semaine les flash arbo.
• Il bénéficie de la « hotline » technique de la chambre.
• Il a participé aux réunions de bilan phytosanitaire et de programme phytosanitaire en hiver 2024-25.
• Son cahier de culture et ses interventions phytosanitaires sont conformes aux réglementations en vigueur, la saisie et la gestion est réalisée sur notre outil de traçabilité SMAG Farmer."""

    y = box_bottom - 1 * cm
    for line in body.splitlines():
        wrapped = wrap(line, width=105)
        for subline in wrapped:
            c.drawString(marge_gauche, y, subline)
            y -= 0.55 * cm
        y -= 0.15 * cm

    # ✅ Signature image full width at bottom
    if signature:
        c.drawImage(signature, x=marge_gauche, y=1.5 * cm,
                    width=largeur - 4 * cm, preserveAspectRatio=True, mask='auto')

    c.save()
    buffer.seek(0)
    return buffer

# --- Streamlit UI ---
st.title("📄 Générateur d'attestations PDF en lot")

uploaded_excel = st.file_uploader("📁 Importer un fichier Excel", type=["xlsx"])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        required_cols = {"Nom", "Date", "Commune", "CodePostal"}
        if not required_cols.issubset(df.columns):
            st.error("❌ Le fichier doit contenir les colonnes : Nom, Date, Commune, CodePostal")
        else:
            st.success("✅ Données chargées, génération en cours...")

            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for index, row in df.iterrows():
                    nom = row["Nom"]
                    date_str = row["Date"].strftime("%d/%m/%Y") if isinstance(row["Date"], (datetime.date, datetime.datetime)) else row["Date"]
                    commune = row["Commune"]
                    code_postal = str(row["CodePostal"])

                    pdf_bytes = generer_pdf(nom, date_str, commune, code_postal, logo_image, signature_image)
                    zip_file.writestr(f"attestation_{nom.replace(' ', '_')}.pdf", pdf_bytes.read())

            zip_buffer.seek(0)
            st.download_button("📥 Télécharger toutes les attestations (.zip)", data=zip_buffer, file_name="attestations.zip", mime="application/zip")

    except Exception as e:
        st.error(f"Erreur lors du traitement : {e}")
