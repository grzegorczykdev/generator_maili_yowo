import streamlit as st
from docxtpl import DocxTemplate, RichText
import io
import zipfile
from pathlib import Path

st.set_page_config(page_title="Generator Maili YoWo", layout="centered")

# Ścieżka do folderu z szablonami
TEMPLATES_DIR = Path(__file__).parent / "templates"

st.title("📄 Generator Maili YoWo")
st.write("Wypełnij poniższe dane, aby wygenerować komplet dokumentów.")

# --- FORMULARZ ---
with st.form("project_form"):
    col1, col2 = st.columns(2)
    
    with col1:
        nazwa_projektu = st.text_input("Nazwa projektu", placeholder="Np. Green Energy")
        typ_projektu = st.selectbox("Typ projektu", ["Youth Exchange", "Training Course"])
        miasto = st.text_input("Miasto")
        kraj = st.text_input("Kraj")
        link_infopack = st.text_input("Link do infopacku")

    with col2:
        data_start = st.date_input("Data rozpoczęcia")
        data_koniec = st.date_input("Data zakończenia")
        deadline_potwierdzenie = st.date_input("Deadline na potwierdzenie udziału")
        dni = st.number_input("Ilość możliwych dodatkowych dni podróży", min_value=1, step=1)
        kwota = st.number_input("Kwota zwrotu kosztów podróży (euro)", min_value=0, step=1)
        imie_nazwisko = st.text_input("Imię i nazwisko do stopki", placeholder="Np. Jan Kowalski")
    
    submit_button = st.form_submit_button("Generuj")

# --- LOGIKA GENEROWANIA ---
if submit_button:
    # Pobieranie plików szablonów z folderu templates (pomijamy ~$... - tymczasowe pliki Word)
    template_files = [f for f in TEMPLATES_DIR.glob("*.docx") if not f.name.startswith("~$")]
    
    if not template_files:
        st.error("W folderze 'templates' nie znaleziono żadnych plików .docx!")
    else:
        # Mapowanie typu projektu na polski tekst
        typ_projektu_pl = {"Youth Exchange": "wymianę młodzieży", "Training Course": "kurs szkoleniowy"}[typ_projektu]
        
        # Słownik bazowy (link_infopack dodamy w pętli jako RichText)
        context_base = {
            'nazwa_projektu': nazwa_projektu,
            'imie_nazwisko': imie_nazwisko,
            'typ_projektu': typ_projektu_pl,
            'data_start': data_start.strftime("%d.%m.%Y"),
            'data_koniec': data_koniec.strftime("%d.%m.%Y"),
            'deadline_potwierdzenie': deadline_potwierdzenie.strftime("%d.%m.%Y"),
            'miasto': miasto,
            'kraj': kraj,
            'kwota': f"{kwota} EUR",
            'dni': dni,
        }

        # Tworzenie pliku ZIP w pamięci
        nazwa_bezpieczna = nazwa_projektu.replace(" ", "_").strip() or "projekt"
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for template_path in template_files:
                try:
                    doc = DocxTemplate(str(template_path))
                    rt_link = RichText()
                    rt_link.add("[LINK]", url_id=doc.build_url_id(link_infopack))
                    context = {**context_base, 'link_infopack': rt_link}
                    doc.render(context)
                    
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    
                    typ_szablonu = template_path.stem.removesuffix("_szablon") if template_path.stem.endswith("_szablon") else template_path.stem
                    nazwa_pliku = f"{nazwa_bezpieczna}_{typ_szablonu}_email.docx"
                    zip_file.writestr(nazwa_pliku, doc_io.getvalue())
                except Exception as e:
                    st.error(f"**Błąd w szablonie:** `{template_path.name}`\n\n{str(e)}")
                    st.info("""
                    **Jak naprawić błąd w szablonie Word:**
                    - Upewnij się, że tagi typu `{{ zmienna }}` lub `{{r link_infopack }}` są zapisane **jako jedna ciągła całość** – bez pogrubienia/cursywy w środku
                    - Sprawdź, czy nie ma pustych tagów `{{ }}` ani niedomkniętych `{{` bez `}}`
                    - Każdy tag musi mieć spacje: `{{ nazwa }}` nie `{{nazwa}}`
                    - Dla linku używaj: `{{r link_infopack }}` (z literą r i spacjami)
                    """)
                    st.stop()

        st.success("✅ Dokumenty gotowe!")
        
        st.download_button(
            label="Pobierz paczkę ZIP",
            data=zip_buffer.getvalue(),
            file_name="dokumentacja_projektu.zip",
            mime="application/zip"
        )