# app.py
import streamlit as st
from processor import generate_final_excel

st.set_page_config(page_title="Stempel vs Telematik – Auswertung", page_icon="📊", layout="centered")

st.title("Stempel vs Telematik – Auswertung (Ein-Klick)")

st.markdown("""
Lade bitte **Stempelprotokoll (PDF)** und **Telematik-Datei (CSV oder Excel)** für **einen Tag** hoch.
Die App liest **'von' → start** und **'bis' → stop** direkt aus der Tabelle im PDF,
matcht die Namen, berechnet die Abweichungen und liefert **eine fertige Excel-Auswertung**.
""")

pdf_file = st.file_uploader("Stempelprotokoll (PDF) hochladen", type=["pdf"])
tele_file = st.file_uploader("Telematik-Datei (CSV oder Excel) hochladen", type=["csv", "xlsx", "xlsm", "xls"])
explicit_date = st.text_input("Optional: Datum erzwingen (TT.MM.JJJJ). Leer lassen für Datum aus dem PDF-Kopf.", value="")

if st.button("Auswertung erstellen", type="primary"):
    if not pdf_file or not tele_file:
        st.error("Bitte **beide** Dateien hochladen (PDF und Telematik).")
    else:
        try:
            final_bytes, final_name = generate_final_excel(
                pdf_bytes=pdf_file.read(),
                tele_bytes=tele_file.read(),
                tele_filename=tele_file.name,
                explicit_date=explicit_date.strip() or None
            )
            st.success("Fertig! Hier kannst du die Auswertung herunterladen.")
            st.download_button(
                "⬇️ Auswertung herunterladen",
                data=final_bytes,
                file_name=final_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erzeugung fehlgeschlagen: {e}")
