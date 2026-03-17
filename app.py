import streamlit as st
from docxtpl import DocxTemplate, RichText
import io
import zipfile
from pathlib import Path

st.set_page_config(page_title="Generator Maili YoWo", layout="centered")

# Ścieżka do folderu z szablonami
TEMPLATES_DIR = Path(__file__).parent / "templates"

# Kraj w miejscowniku (dla "w X"): nazwa -> forma w języku polskim
KRAJ_MIEJSCOWNIK = {
    "albania": "Albanii", "albanii": "Albanii",
    "andora": "Andorze",
    "armenia": "Armenii",
    "austria": "Austrii",
    "azersbajdżan": "Azerbejdżanie", "azerbejdżan": "Azerbejdżanie", "azerbaijan": "Azerbejdżanie",
    "belgia": "Belgii", "belgium": "Belgii",
    "białoruś": "Białorusi", "bialorus": "Białorusi", "belarus": "Białorusi",
    "bośnia": "Bośni", "bośnia i hercegowina": "Bośni i Hercegowinie", "bosnia": "Bośni",
    "bułgaria": "Bułgarii", "bulgaria": "Bułgarii",
    "chorwacja": "Chorwacji", "croatia": "Chorwacji",
    "cypr": "Cyprze", "cyprus": "Cyprze",
    "czarnogóra": "Czarnogórze", "montenegro": "Czarnogórze",
    "czechy": "Czechach", "czech republic": "Czechach", "česko": "Czechach",
    "dania": "Danii", "denmark": "Danii",
    "estonia": "Estonii",
    "finlandia": "Finlandii", "finland": "Finlandii",
    "francja": "Francji", "france": "Francji",
    "grecja": "Grecji", "greece": "Grecji",
    "gruzja": "Gruzji", "georgia": "Gruzji",
    "hiszpania": "Hiszpanii", "spain": "Hiszpanii",
    "holandia": "Holandii", "niderlandy": "Niderlandach", "netherlands": "Holandii",
    "irlandia": "Irlandii", "ireland": "Irlandii",
    "islandia": "Islandii", "iceland": "Islandii",
    "kosowo": "Kosowie", "kosovo": "Kosowie",
    "liechtenstein": "Liechtensteinie",
    "litwa": "Litwie", "lithuania": "Litwie",
    "luksemburg": "Luksemburgu", "luxembourg": "Luksemburgu",
    "łotwa": "Łotwie", "lotwa": "Łotwie", "latvia": "Łotwie",
    "macedonia": "Macedonii", "macedonia północna": "Macedonii Północnej", "north macedonia": "Macedonii Północnej",
    "malta": "Malcie", "malta": "Malcie",
    "mołdawia": "Mołdawii", "molawia": "Mołdawii", "moldova": "Mołdawii",
    "monako": "Monako", "monaco": "Monako",
    "niemcy": "Niemczech", "germany": "Niemczech", "deutschland": "Niemczech",
    "norwegia": "Norwegii", "norway": "Norwegii",
    "polska": "Polsce", "poland": "Polsce",
    "portugalia": "Portugalii", "portugal": "Portugalii",
    "rumunia": "Rumunii", "romania": "Rumunii",
    "san marino": "San Marino",
    "serbia": "Serbii",
    "słowacja": "Słowacji", "slowacja": "Słowacji", "slovakia": "Słowacji",
    "słowenia": "Słowenii", "slowenia": "Słowenii", "slovenia": "Słowenii",
    "szwecja": "Szwecji", "sweden": "Szwecji",
    "szwajcaria": "Szwajcarii", "switzerland": "Szwajcarii",
    "turcja": "Turcji", "turkey": "Turcji",
    "ukraina": "Ukrainie", "ukraine": "Ukrainie",
    "watykan": "Watykanie", "vatican": "Watykanie",
    "węgry": "Węgrzech", "wegry": "Węgrzech", "hungary": "Węgrzech",
    "wielka brytania": "Wielkiej Brytanii", "wielkiej brytanii": "Wielkiej Brytanii",
    "uk": "Wielkiej Brytanii", "great britain": "Wielkiej Brytanii", "united kingdom": "Wielkiej Brytanii",
    "włochy": "Włoszech", "wlochy": "Włoszech", "italy": "Włoszech", "italia": "Włoszech",
}

st.title("🌻 Generator Maili YoWo")
st.write("Wypełnij poniższe dane, aby wygenerować szablony wiadomości :)")

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
        kwota = st.number_input("Kwota zwrotu kosztów podróży (euro)", min_value=0, step=1, format="%d")
        imie_nazwisko = st.text_input("Imię i nazwisko do stopki", placeholder="Np. Jan Kowalski")
    
    st.markdown("**Dodatkowe dni podróży**")
    dni = st.number_input("Ilość dodatkowych dni", min_value=0, step=1)
    nie_uwzgledniono_dni = st.checkbox("Nie ma w infopacku (nie podajemy informacji w e-mailu)", key="nie_uwzgledniono")
    
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
        typ_projektu_2 = {"Youth Exchange": "wymianą młodzieży", "Training Course": "kursem szkoleniowym"}[typ_projektu]
        ktora_ktory = {"Youth Exchange": "która", "Training Course": "który"}[typ_projektu]  # wymiana (ż) -> która, kurs (m) -> który
        
        # Kraj w miejscowniku (np. Czechy->Czechach dla "w Czechach")
        kraj_normalized = kraj.strip().lower() if kraj else ""
        kraj_2 = KRAJ_MIEJSCOWNIK.get(kraj_normalized, kraj)  # fallback: oryginalna wartość
        
        # Tekst o dodatkowych dniach: pełna formułka albo pusty string
        if nie_uwzgledniono_dni or dni == 0:
            dni_info = ""
        else:
            dni_info = f"🌻 Możesz wykorzystać dodatkowe {dni} dni przed lub po projekcie i zwiedzać, jednak w te extra dni koszt zakwaterowania i jedzenia nie jest zwracany."
        
        # Tematy e-maili w zależności od typu projektu
        temat_zakwalifikowany = f'Erasmus+ {typ_projektu} "[{nazwa_projektu}]" - zaproszenie na projekt'
        temat_odrzucony = f'Erasmus+ {typ_projektu} "[{nazwa_projektu}]" - dziękujemy za zgłoszenie'
        temat_rezerwowy = f'Erasmus+ {typ_projektu} "[{nazwa_projektu}]" - lista rezerwowa'
        
        # Słownik bazowy (link_infopack dodamy w pętli jako RichText)
        context_base = {
            'nazwa_projektu': nazwa_projektu,
            'imie_nazwisko': imie_nazwisko,
            'typ_projektu': typ_projektu_pl,
            'typ_projektu_2': typ_projektu_2,
            'data_start': data_start.strftime("%d.%m.%Y"),
            'data_koniec': data_koniec.strftime("%d.%m.%Y"),
            'deadline_potwierdzenie': deadline_potwierdzenie.strftime("%d.%m.%Y"),
            'miasto': miasto,
            'kraj': kraj,
            'kraj_2': kraj_2,
            'kwota': f"{kwota}",
            'dni': dni,
            'dni_info': dni_info,
            'temat_zakwalifikowany': temat_zakwalifikowany,
            'temat_odrzucony': temat_odrzucony,
            'temat_rezerwowy': temat_rezerwowy,
            'ktora_ktory': ktora_ktory,
        }

        # Tworzenie pliku ZIP w pamięci
        nazwa_bezpieczna = nazwa_projektu.replace(" ", "_").strip() or "projekt"
        
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for template_path in template_files:
                try:
                    doc = DocxTemplate(str(template_path))
                    rt_link = RichText()
                    rt_link.add("[LINK]", url_id=doc.build_url_id(link_infopack), bold=True)
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
            file_name=f"{nazwa_bezpieczna}_maile.zip",
            mime="application/zip"
        )