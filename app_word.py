import streamlit as st
import io
import re
from docx import Document
from docx.shared import Inches

st.set_page_config(page_title="Compilatore Word Completo", page_icon="📝", layout="centered")
st.title("📝 Compilatore Word per Smartphone")
st.write("I campi qui sotto mostrano l'intero testo presente a sinistra di ogni segnaposto.")

# --- 1. CARICAMENTO FILE ---
file_word = st.file_uploader("1. Carica il file Word (.docx)", type=["docx"])
file_firma = st.file_uploader("2. Carica la firma (.png, .jpg)", type=["png", "jpg", "jpeg"])

def estrai_campi_con_tutto_il_testo(doc):
    """Trova i segnaposto [...] e ricava l'etichetta prendendo TUTTO il testo a sinistra"""
    campi_rilevati = {} # Formato: {"[segnaposto]": "Intera frase a sinistra"}
    pattern = r"\[([^\]]+)\]"

    # Funzione interna per estrarre tutto il testo a sinistra
    def ricava_etichetta_completa(testo, segnaposto_intero):
        if segnaposto_intero in testo:
            # Divide la riga in base al segnaposto e prende la prima parte (tutto ciò che c'è a sinistra)
            testo_precedente = testo.split(segnaposto_intero)[0]
            # Pulisce gli spazi bianchi all'inizio e alla fine, e rimuove i due punti o trattini finali
            etichetta = testo_precedente.strip().rstrip(":-_ \t")
            # Se la riga a sinistra non era vuota, la restituisce
            if len(etichetta) > 0:
                return etichetta
        return None

    # Scansione dei paragrafi normali
    for p in doc.paragraphs:
        trovati = re.findall(pattern, p.text)
        for t in trovati:
            segnaposto = f"[{t}]"
            etichetta_completa = ricava_etichetta_completa(p.text, segnaposto)
            # Se a sinistra è vuoto (es. il segnaposto è a inizio riga), usa il testo dentro le parentesi
            campi_rilevati[segnaposto] = etichetta_completa if etichetta_completa else t

    # Scansione delle tabelle
    for tabella in doc.tables:
        for riga in tabella.rows:
            for cella in riga.cells:
                for p in cella.paragraphs:
                    trovati = re.findall(pattern, p.text)
                    for t in trovati:
                        segnaposto = f"[{t}]"
                        etichetta_completa = ricava_etichetta_completa(p.text, segnaposto)
                        campi_rilevati[segnaposto] = etichetta_completa if etichetta_completa else t

    return campi_rilevati

if file_word:
    doc_bytes = io.BytesIO(file_word.read())
    doc = Document(doc_bytes)
    
    # Estraiamo i campi abbinati a tutto il testo che li precede
    mappa_campi = estrai_campi_con_tutto_il_testo(doc)
    
    if mappa_campi:
        st.subheader("✍️ Compila il documento:")
        dati_da_inviare = {}
        
        # Genera le caselle di testo mostrando l'intera frase a sinistra
        for segnaposto, testo_sinistra in mappa_campi.items():
            risposta_utente = st.text_input(
                label=f"{testo_sinistra}:", 
                key=segnaposto,
                placeholder=f"Inserisci qui..."
            )
            dati_da_inviare[segnaposto] = risposta_utente
            
        # --- GENERAZIONE ---
        if st.button("🚀 Genera e Firma Documento", type="primary"):
            # Sostituzione nei paragrafi
            for p in doc.paragraphs:
                for k, v in dati_da_inviare.items():
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        
            # Sostituzione nelle tabelle
            for tabella in doc.tables:
                for riga in tabella.rows:
                    for cella in riga.cells:
                        for p in cella.paragraphs:
                            for k, v in dati_da_inviare.items():
                                if k in p.text:
                                    p.text = p.text.replace(k, v)
                                    
            # Inserimento Firma
            if file_firma:
                firma_bytes = io.BytesIO(file_firma.read())
                firma_inserita = False
                for p in doc.paragraphs:
                    if "Firma:" in p.text:
                        p_firma = p.insert_paragraph_before()
                        p_firma.add_run().add_picture(firma_bytes, width=Inches(2.0))
                        firma_inserita = True
                        break
                if not firma_inserita:
                    doc.add_picture(firma_bytes, width=Inches(2.0))
            
            # Preparazione download
            output = io.BytesIO()
            doc.save(output)
            output.seek(0)
            
            st.success("🎉 Documento pronto per il download!")
            st.download_button(
                label="📥 Scarica il file compilato",
                data=output,
                file_name="documento_compilato.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("⚠️ Nessun segnaposto [...] trovato nel testo o nelle tabelle del documento.")
