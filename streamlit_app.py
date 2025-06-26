import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
import io
import zipfile
import tempfile

st.set_page_config(page_title="Generatore Attestati", layout="centered")
st.title("🎓 Generatore Attestati 2025")

with st.form("upload_form"):
    pdf_file = st.file_uploader("📄 PDF modello (senza campi)", type=["pdf"])
    excel_file = st.file_uploader("📊 File Excel (.xlsx)", type=["xlsx"])
    bold_font = st.file_uploader("🔡 Font Exo2-Bold (.ttf)", type=["ttf"])
    light_font = st.file_uploader("🔡 Font Exo2-Light (.ttf)", type=["ttf"])
    submitted = st.form_submit_button("Genera Attestati")

if submitted:
    if not all([pdf_file, excel_file, bold_font, light_font]):
        st.error("⚠️ Devi caricare tutti i file richiesti.")
    else:
        with st.spinner("🎯 Generazione in corso..."):

            # Registra font
            pdfmetrics.registerFont(TTFont("Exo2-ExtraBold", bold_font))
            pdfmetrics.registerFont(TTFont("Exo2-Light", light_font))

            # Leggi dati
            df = pd.read_excel(excel_file)
            assert {'NOME', 'COGNOME'}.issubset(df.columns), "Excel deve contenere colonne NOME e COGNOME"

            # Coordinate fisse
            coord_nome   = (176, 284)
            coord_titolo = (266, 253)
            coord_data   = (176, 196)

            data_evento = "mercoledì 18 giugno 2025"
            titolo = "NIS2: le misure di sicurezza"

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for _, row in df.iterrows():
                    nome = str(row["NOME"]).strip()
                    cognome = str(row["COGNOME"]).strip()
                    nome_completo = f"{nome} {cognome}"

                    # Leggi PDF
                    base_pdf = PdfReader(pdf_file)
                    w = float(base_pdf.pages[0].mediabox.width)
                    h = float(base_pdf.pages[0].mediabox.height)

                    # Crea overlay
                    packet = io.BytesIO()
                    c = canvas.Canvas(packet, pagesize=(w, h))
                    c.setFont("Exo2-Bold", 24)
                    c.drawString(*coord_nome, nome_completo)
                    c.setFont("Exo2-Light", 14)
                    c.drawString(*coord_titolo, titolo)
                    c.drawString(*coord_data, data_evento)
                    c.save()
                    packet.seek(0)

                    overlay = PdfReader(packet)
                    output = PdfWriter()

                    page = base_pdf.pages[0]
                    page.merge_page(overlay.pages[0])
                    output.add_page(page)

                    for p in base_pdf.pages[1:]:
                        output.add_page(p)

                    out_pdf = io.BytesIO()
                    output.write(out_pdf)
                    out_pdf.seek(0)

                    filename = f"attestato_18giugno2025_{cognome}.pdf"
                    zipf.writestr(filename, out_pdf.read())

            st.success("✅ Attestati generati con successo!")

            st.download_button(
                label="📥 Scarica archivio ZIP",
                data=zip_buffer.getvalue(),
                file_name="attestati_18giugno2025.zip",
                mime="application/zip"
            )
