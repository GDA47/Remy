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

# --- Format date in French style
def formater_date_lettres(date_str):
    date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y")
    mois = ["janvier", "février", "mars", "avril", "mai", "juin",
            "juillet", "août", "septembre", "octobre", "novembre", "décembre"]
    return f"{date_obj.day} {mois[date_obj.month - 1]} {date_obj.year}"

# --- Generate PDF
def generer_pdf(nom, date_str, commune, code_postal, logo, signature):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largeur, hauteur = A4

    marge_gauche = 2 * cm

    # Logo full width
    if logo:
        c.drawImage(logo, x=marge_gauche, y=hauteur - 6 * cm,
                    width=largeur - 4 * cm, preserveAspectRatio=True, mask='auto')

    # Date
    c.setFont("Helvetica", 12)
    c.setFillColorRGB(0, 0, 0)
    c.drawString(marge_gauche, hauteur - 6.2 * cm, f"Agen, le {formater_date_lettres(date_str)}")

    # Title box
    box_top = hauteur - 7.5 * cm
    box_bottom = box_top - 2.4 * cm
    c.setStrokeColorRGB(0.3, 0.6, 0.3)
    c.setFillColorRGB(0.85, 0.95, 0.85)
    c.rect(marge_gauche, box_bottom, largeur - 4 * cm, 2.4 * cm, fill=1, stroke=1)

    c.setFont("Helvetica-Bold", 14)
    c.setFillColorRGB(0, 0, 0)
    c.drawCentredString(largeur / 2, box_top - 0.9 * cm, "Attestation de Suivi Technique")
    c.drawCentredString(largeur / 2, box_top - 1.7 * cm, "Pomme Production Fruitière Intégrée")

    # Body
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

    # Signature
    if signature:
        c.drawImage(signature, x=marge_gauche, y=1.5 * cm,
                    width=largeur - 4 * cm, preserveAspectRatio=True, mask='auto')

    c.save()
    buffer.seek(0)
    return buffer

# --- UI ---
st.title("📄 Générateur d'attestations PDF")

with st.expander("ℹ️ Instructions pour le fichier Excel requis"):
    st.markdown("""
    Le fichier Excel doit contenir **une ligne par attestation** avec les colonnes suivantes :

    | Nom                 | Date        | Commune             | CodePostal |
    |----------------------|-------------|----------------------|------------|
    | Mme Alain LOUBIERES  | 27/01/2025  | CLERMONT-DESSOUS     | 47130      |

    **⚠️ Astuces :**
    - Le format de la date doit être `JJ/MM/AAAA`
    - La colonne `CodePostal` doit être un nombre ou une chaîne à 5 chiffres
    """)

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
        st.error(f"❌ Erreur lors du traitement du fichier : {e}")

# --- Manual Entry ---
st.markdown("---")
st.subheader("📝 Générer une attestation manuellement")

with st.form("manual_form"):
    nom_manual = st.text_input("Nom complet")
    date_manual = st.date_input("Date de l'attestation", datetime.date.today())
    commune_manual = st.text_input("Commune")
    cp_manual = st.text_input("Code postal")

    submitted = st.form_submit_button("📄 Générer l'attestation")

if submitted:
    if not (nom_manual and commune_manual and cp_manual):
        st.warning("Veuillez remplir tous les champs pour générer l’attestation.")
    else:
        try:
            date_str_manual = date_manual.strftime("%d/%m/%Y")
            pdf_buffer = generer_pdf(nom_manual, date_str_manual, commune_manual, cp_manual, logo_image, signature_image)
            st.success("✅ Attestation générée avec succès")
            st.download_button(
                label="📥 Télécharger l'attestation",
                data=pdf_buffer,
                file_name=f"attestation_{nom_manual.replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"❌ Une erreur s’est produite : {e}")
