import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfReader, PdfWriter
import io
import zipfile
from pathlib import Path

st.set_page_config(page_title="Generatore Attestati 2025", layout="centered")
st.title("🎓 Generatore Attestati - 18 Giugno 2025")

# Percorsi font nella cartella del repo
FONT_BOLD_PATH = Path("fonts/Exo2-Bold.ttf")
FONT_LIGHT_PATH = Path("fonts/Exo2-Light.ttf")

# Registra font al caricamento dell'app
pdfmetrics.registerFont(TTFont("Exo2-Bold", str(FONT_BOLD_PATH)))
pdfmetrics.registerFont(TTFont("Exo2-Light", str(FONT_LIGHT_PATH)))

with st.form("upload_form"):
    pdf_file = st.file_uploader("📄 PDF modello (senza campi)", type=["pdf"])
    xls_file = st.file_uploader("📊 File Excel (.xlsx)", type=["xlsx"])
    submitted = st.form_submit_button("Genera Attestati")

if submitted:
    if not all([pdf_file, xls_file]):
        st.error("⚠️ Carica sia il PDF modello sia il file Excel.")
    else:
        with st.spinner("✍️ Generazione in corso..."):

            # Leggi Excel
            df = pd.read_excel(xls_file)
            assert {'NOME', 'COGNOME'}.issubset(df.columns), "Il file Excel deve avere colonne NOME e COGNOME"

            # Coordinate e testi fissi
            coord_nome   = (176, 284)
            coord_titolo = (266, 253)
            coord_data   = (176, 196)
            titolo_txt   = "NIS2: le misure di sicurezza"
            data_txt     = "mercoledì 18 giugno 2025"

            # Output ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zipf:
                for _, r in df.iterrows():
                    nome_cogn = f"{r['NOME'].strip()} {r['COGNOME'].strip()}"
                    pdf_reader = PdfReader(pdf_file)
                    page = pdf_reader.pages[0]
                    w, h = float(page.mediabox.width), float(page.mediabox.height)

                    # Crea overlay
                    packet = io.BytesIO()
                    c = canvas.Canvas(packet, pagesize=(w, h))
                    c.setFont("Exo2-Bold", 24)
                    c.drawString(*coord_nome, nome_cogn)
                    c.setFont("Exo2-Light", 18)
                    c.drawString(*coord_titolo, titolo_txt)
                    c.drawString(*coord_data, data_txt)
                    c.save()
                    packet.seek(0)

                    overlay = PdfReader(packet)
                    writer = PdfWriter()
                    page.merge_page(overlay.pages[0])
                    writer.add_page(page)

                    for extra_page in pdf_reader.pages[1:]:
                        writer.add_page(extra_page)

                    pdf_out = io.BytesIO()
                    writer.write(pdf_out)
                    pdf_out.seek(0)

                    filename = f"attestato_18giugno2025_{r['COGNOME'].strip()}.pdf"
                    zipf.writestr(filename, pdf_out.read())

            st.success("✅ Attestati generati con successo!")

            st.download_button(
                label="📥 Scarica ZIP",
                data=zip_buffer.getvalue(),
                file_name="attestati_18giugno2025.zip",
                mime="application/zip"
            )
