import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime, time, timedelta
import os
import urllib.parse

st.set_page_config(page_title="Wyszukiwarka Zastępstw", page_icon="📋", layout="wide")
st.title("Wyszukiwarka Zastępstw 📋")

# --- ZARZĄDZANIE PAMIĘCIĄ SESJI I CALLBACKI ---
if "search_results" not in st.session_state:
    st.session_state.search_results = None
    st.session_state.teacher_counts = {}

def update_end_time(group_idx):
    start_key = f"start_{group_idx}"
    end_key = f"end_{group_idx}"
    if start_key in st.session_state:
        start_val = st.session_state[start_key]
        start_dt = datetime.combine(datetime.today(), start_val)
        st.session_state[end_key] = (start_dt + timedelta(minutes=90)).time()

def update_branches():
    if "branch_0" in st.session_state:
        new_branch = st.session_state["branch_0"]
        for i in range(1, 6):
            if f"branch_{i}" in st.session_state:
                st.session_state[f"branch_{i}"] = new_branch

def next_tpl(idx):
    st.session_state[f"tpl_{idx}"] = (st.session_state.get(f"tpl_{idx}", 0) + 1) % 6

# --- FUNKCJE I SŁOWNIKI BAZOWE ---
def is_cell_free(cell):
    if not cell.fill or cell.fill.fill_type is None: return True
    color = getattr(cell.fill.start_color, "index", None)
    if color in [0, "00000000", "FFFFFFFF", "None", None]: return True
    rgb = getattr(cell.fill.start_color, "rgb", None)
    if rgb in ["00000000", "FFFFFFFF", None]: return True
    return False

days = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek"]
levels = ["3-5 LAT", "0", "1", "2", "3", "4", "5", "6", "7", "8+", "MASTER"]
day_mapping = {"Poniedziałek": "PON", "Wtorek": "WT", "Środa": "ŚR", "Czwartek": "CZW", "Piątek": "PT"}
branches = {
    "-": None,
    "Ursus 1 (U1)": "U1", "Ursus 2 (U2)": "U2", "Komorów (K)": "K",
    "Michałowice (M)": "M", "Nowa Wieś (NW)": "NW", "Pruszków (P)": "P"
}
loc_grammar = {
    "U1": "Ursusa 1", "U2": "Ursusa 2", "K": "Komorowa", 
    "M": "Michałowic", "NW": "Nowej Wsi", "P": "Pruszkowa"
}

# --- OBSŁUGA PLIKU ---
default_file = "Zastępstwa 2026_27 - KUM&Co.xlsx"

st.info("💡 **Wskazówka:** Aplikacja domyślnie korzysta z zapisanego na serwerze grafiku. Jeśli wiesz, że grafik uległ niedawno zmianie, warto wgrać jego zaktualizowaną wersję poniżej.")
uploaded_file = st.file_uploader("Wgraj zaktualizowany grafik (opcjonalnie)", type=["xlsx"])

if uploaded_file is not None:
    active_file = uploaded_file
    st.success("✅ Używam wgranego przez Ciebie nowszego pliku!")
elif os.path.exists(default_file):
    active_file = default_file
    st.success(f"✅ Używam domyślnego pliku z serwera: **{default_file}**.")
else:
    active_file = None
    st.warning("⚠️ Nie znaleziono domyślnego pliku na serwerze. Proszę wgrać grafik ręcznie.")

# --- SPRAWDZANIE WŁASNEGO PLANU ---
st.markdown("---")
with st.expander("📅 Sprawdź swój własny plan (opcjonalnie)"):
    st.write("Szybki podgląd Twoich dzisiejszych zajęć, żeby łatwiej było zaplanować ewentualną zamianę z innym lektorem.")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        my_name_input = st.text_input("Podaj swoje imię i nazwisko", value="").strip()
    with col_m2:
        my_day_input = st.selectbox("Wybierz dzień", days, key="my_day")
        
    if st.button("Pokaż mój plan"):
        if not active_file:
            st.error("Brak pliku grafiku.")
        elif not my_name_input:
            st.warning("Wpisz imię i nazwisko.")
        else:
            try:
                wb_temp = openpyxl.load_workbook(active_file, data_only=True)
                sheet_my = wb_temp[day_mapping[my_day_input]]
                header_my = [(cell.column, cell.value) for cell in sheet_my[4] if isinstance(cell.value, time)]
                
                found_my_row = None
                for r in range(4, sheet_my.max_row + 1):
                    v = str(sheet_my.cell(r, 1).value).strip()
                    if v != "None" and my_name_input.lower() in v.lower():
                        found_my_row = r
                        break
                        
                if found_my_row:
                    my_busy = []
                    for c, t_val in header_my:
                        if not is_cell_free(sheet_my.cell(found_my_row, c)):
                            my_busy.append(t_val)
                            
                    if my_busy:
                        start_str = my_busy[0].strftime('%H:%M')
                        last_dt = datetime.combine(datetime.today(), my_busy[-1]) + timedelta(minutes=15)
                        end_str = last_dt.strftime('%H:%M')
                        st.success(f"W {my_day_input.lower()} masz zajęcia **od {start_str} do {end_str}**.")
                    else:
                        st.success(f"W {my_day_input.lower()} nie masz w grafiku żadnych zajęć (masz wolne!).")
                else:
                    st.error(f"Nie znaleziono osoby: {my_name_input} w grafiku na {my_day_input.lower()}.")
            except Exception as e:
                st.error(f"Błąd sprawdzania planu: {e}")

# --- USTAWIENIA NA GŁÓWNYM EKRANIE ---
st.markdown("---")
st.write("### Opcje wyszukiwania")
col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    exclude_me = st.text_input("Nie pokazuj mojego emaila", placeholder="Wpisz swoje imię i nazwisko, np. Jan Kowalski").strip().lower()
with col_opt2:
    st.write("") 
    is_multiple = st.checkbox("Szukam zastępstw dla więcej niż 1 grupy", value=False)
    if is_multiple:
        num_groups = st.number_input("Ile grup?", min_value=2, max_value=5, value=2)
    else:
        num_groups = 1

st.markdown("---")

# --- KONFIGURACJA GRUP ---
groups_config = []
st.write("### Parametry Grupy 1")

if "start_0" not in st.session_state:
    st.session_state["start_0"] = time(15, 50)
    st.session_state["end_0"] = time(17, 20)

col1, col2, col3 = st.columns(3)
with col1: default_day = st.selectbox("Dzień", days, key="day_0")
with col2: default_level = st.selectbox("Poziom", levels, key="level_0")
with col3: default_branch = st.selectbox("Która filia?", list(branches.keys()), key="branch_0", on_change=update_branches)

col_t1, col_t2 = st.columns(2)
with col_t1: 
    default_start = st.time_input("Czas rozpoczęcia", key="start_0", on_change=update_end_time, args=(0,))
with col_t2: 
    default_end = st.time_input("Czas zakończenia", key="end_0")

groups_config.append({
    "day": default_day, "level": default_level, "branch": default_branch, 
    "start": default_start, "end": default_end
})

if is_multiple:
    for i in range(1, int(num_groups)):
        if f"start_{i}" not in st.session_state:
            st.session_state[f"start_{i}"] = st.session_state["start_0"]
            st.session_state[f"end_{i}"] = st.session_state["end_0"]
        if f"branch_{i}" not in st.session_state:
            st.session_state[f"branch_{i}"] = st.session_state.get("branch_0", list(branches.keys())[0])

        st.markdown("---")
        st.write(f"### Parametry Grupy {i+1}")
        c1, c2, c3 = st.columns(3)
        with c1: g_day = st.selectbox("Dzień", days, index=days.index(default_day), key=f"day_{i}")
        with c2: g_level = st.selectbox("Poziom", levels, key=f"level_{i}")
        with c3: g_branch = st.selectbox("Która filia?", list(branches.keys()), key=f"branch_{i}")
        
        ct1, ct2 = st.columns(2)
        with ct1: 
            g_start = st.time_input("Czas rozpoczęcia", key=f"start_{i}", on_change=update_end_time, args=(i,))
        with ct2: 
            g_end = st.time_input("Czas zakończenia", key=f"end_{i}")
        
        groups_config.append({
            "day": g_day, "level": g_level, "branch": g_branch, 
            "start": g_start, "end": g_end
        })

# --- FUNKCJE POMOCNICZE CZĘŚĆ 2 ---
def overlaps(cell_time, req_start, req_end):
    return req_start < cell_time < req_end

def get_commute_warning(teacher_row, day_sheet, req_start, target_branch_code, teacher_branch_str):
    if not target_branch_code: return ""
    t_branch = teacher_branch_str.upper().replace(" ", "").replace("&", "+").split("+")[0]
    if t_branch not in loc_grammar or target_branch_code not in loc_grammar: return ""
    if t_branch == target_branch_code: return ""
    
    last_busy_end = None
    dummy = datetime(2000, 1, 1)
    req_dt = datetime.combine(dummy, req_start)
    
    for cell in day_sheet[4]:
        if isinstance(cell.value, time):
            col_dt = datetime.combine(dummy, cell.value)
            if col_dt <= req_dt:
                t_cell = day_sheet.cell(row=teacher_row, column=cell.column)
                if not is_cell_free(t_cell):
                    last_busy_end = col_dt 
                    
    if last_busy_end and last_busy_end <= req_dt:
        diff_mins = int((req_dt - last_busy_end).total_seconds() / 60)
        if 0 <= diff_mins <= 60:
            od = loc_grammar[t_branch]
            return f"Kończę o {last_busy_end.strftime('%H:%M')} ({diff_mins} min z: {od})"
    return ""

st.markdown("---")

# --- GŁÓWNY PROCES ---
if st.button("Znajdź Zastępstwo 🚀", use_container_width=True):
    if active_file is None:
        st.error("⚠ Proszę najpierw wgrać plik z grafikiem!")
    else:
        try:
            wb = openpyxl.load_workbook(active_file, data_only=True)
            if not isinstance(active_file, str): 
                active_file.seek(0)
                
            df_levels = pd.read_excel(active_file, sheet_name="Lektorzy i poziomy grup na zast", header=None)
            
            all_group_results = []
            
            for g_idx, g_conf in enumerate(groups_config):
                level_col = None
                
                for col_idx in df_levels.columns:
                    if str(df_levels.iloc[4][col_idx]).strip() == str(g_conf["level"]).strip():
                        level_col = col_idx
                        break
                        
                q_teachers = []
                if level_col is not None:
                    for idx, row in df_levels.iterrows():
                        if idx < 4: continue
                        val = str(row[level_col]).strip().lower()
                        if val == "v":
                            t_name = str(row[1]).strip()
                            if exclude_me and exclude_me in t_name.lower():
                                continue 
                            b_val = str(row[2]).strip()
                            phone_val = str(row[3]).strip()
                            if phone_val == "nan": phone_val = "-"
                            email = str(row[4]).strip()
                            
                            if t_name != "nan":
                                q_teachers.append({"name": t_name, "email": email, "phone": phone_val, "branch": b_val})
                                
                excel_day = day_mapping[g_conf["day"]]
                day_sheet = wb[excel_day]
                
                overlapping_cols = {}
                for cell in day_sheet[4]:
                    if isinstance(cell.value, time):
                        if overlaps(cell.value, g_conf["start"], g_conf["end"]):
                            overlapping_cols[cell.column] = cell.value
                            
                a_teachers = []
                if overlapping_cols:
                    for t in q_teachers:
                        t_row = None
                        for row in range(4, day_sheet.max_row + 1):
                            c_val = str(day_sheet.cell(row=row, column=1).value).strip()
                            if c_val != "None" and (t["name"] in c_val or c_val in t["name"]):
                                t_row = row
                                break
                                
                        if t_row:
                            is_free = True
                            for col in overlapping_cols.keys():
                                if not is_cell_free(day_sheet.cell(row=t_row, column=col)):
                                    is_free = False
                                    break
                                    
                            if is_free:
                                target_b_code = branches[g_conf["branch"]]
                                commute_warn = get_commute_warning(t_row, day_sheet, g_conf["start"], target_b_code, t["branch"])
                                
                                is_at_branch = False
                                if target_b_code:
                                    b_upper = t["branch"].upper()
                                    if "KUM" in b_upper and "CO" in b_upper:
                                        is_at_branch = True
                                    else:
                                        t_branches = [b.strip() for b in b_upper.replace(" ", "").replace("&", "+").split("+")]
                                        if target_b_code in t_branches:
                                            is_at_branch = True
                                            
                                a_teachers.append({
                                    "Nauczyciel": t["name"],
                                    "E-mail": t["email"],
                                    "Telefon": t["phone"],
                                    "Filie": t["branch"],
                                    "_is_at_branch": is_at_branch,
                                    "_commute_warn": commute_warn
                                })
                                
                all_group_results.append(a_teachers)
                
            counts = {}
            for res in all_group_results:
                for t in res:
                    counts[t["Nauczyciel"]] = counts.get(t["Nauczyciel"], 0) + 1
                    
            st.session_state.search_results = all_group_results
            st.session_state.teacher_counts = counts
            
        except Exception as e:
            st.error(f"Wystąpił błąd: {e}")

# --- WYSWIETLANIE I FILTROWANIE WYNIKÓW ---
if st.session_state.search_results is not None:
    st.write("### ⚙️ Filtry wyników")
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    with col_f1:
        show_only_all = st.toggle("Tylko lektorzy na WSZYSTKIE grupy") if is_multiple else False
    with col_f2:
        exclude_bad_commute = st.toggle("Wyklucz lektorów z małą ilością czasu na dojazd")
    with col_f3:
        only_same_branch = st.toggle("Tylko lektorzy uczący dziś w tej filii")
    with col_f4:
        show_phones = st.toggle("Pokaż numery telefonów")

    st.info("💡 **Legenda (najedź na ikonę):** 🔥 - Wszystkie grupy | ⭐ - Część grup | 🏫 - Ta sama filia | ⚠️ - Problem z dojazdem")

    for g_idx, (g_conf, res) in enumerate(zip(groups_config, st.session_state.search_results)):
        st.markdown(f"## Wyniki dla Grupy {g_idx+1}")
        st.write(f"**{g_conf['day']} | {g_conf['start'].strftime('%H:%M')} - {g_conf['end'].strftime('%H:%M')} | Poziom: {g_conf['level']} | Filia: {g_conf['branch']}**")
        
        display_data = []
        for t in res:
            count = st.session_state.teacher_counts[t["Nauczyciel"]]
            
            if show_only_all and count < int(num_groups): continue
            if exclude_bad_commute and t["_commute_warn"] != "": continue
            if only_same_branch and not t["_is_at_branch"]: continue
                
            notes_html = []
            if t["_is_at_branch"]:
                notes_html.append('<span title="Uczę w tej filii :)">🏫</span>')
            if t["_commute_warn"]:
                notes_html.append(f'<span title="{t["_commute_warn"]}">⚠️</span>')
                
            sort_all = False
            if is_multiple:
                odmiana = "zastępstwa" if count in [2, 3, 4] else "zastępstw"
                if count == int(num_groups):
                    notes_html.append(f'<span title="Mogę wziąć WSZYSTKIE {count} {odmiana}!">🔥</span>')
                    sort_all = True
                elif count > 1:
                    notes_html.append(f'<span title="Mogę wziąć {count} {odmiana}.">⭐</span>')
                    
            notes_str = " ".join(notes_html) if notes_html else "-"
            
            display_data.append({
                "Nauczyciel": t["Nauczyciel"],
                "E-mail": t["E-mail"],
                "Telefon": t["Telefon"],
                "Notatki": notes_str,
                # Klucze do sortowania (nie wyświetlają się)
                "_sort_all": sort_all,
                "_sort_branch": t["_is_at_branch"],
                "_sort_warn": t["_commute_warn"] == "",
                "_sort_count": count
            })
            
        if not display_data:
            st.warning("Brak nauczycieli spełniających wybrane kryteria i filtry.")
        else:
            # Sortowanie: najlepsze dopasowania trafiają na górę tabeli
            display_data.sort(key=lambda x: (
                not x["_sort_all"], 
                not x["_sort_branch"], 
                not x["_sort_warn"], 
                -x["_sort_count"], 
                x["Nauczyciel"]
            ))

            emails = [d["E-mail"] for d in display_data if d["E-mail"] != "nan" and "@" in d["E-mail"]]
            
            st.info("💡 **Pamiętaj, żeby wpisać się w tabelkę i załączyć w DW lidera, biuro i metodyków swojej filii :)**")
            st.code("; ".join(emails), language="text")
            
            if show_phones:
                md_table = "| Nauczyciel | E-mail | Telefon | Notatki |\n|---|---|---|---|\n"
                for d in display_data:
                    md_table += f"| {d['Nauczyciel']} | {d['E-mail']} | {d['Telefon']} | {d['Notatki']} |\n"
            else:
                md_table = "| Nauczyciel | E-mail | Notatki |\n|---|---|---|\n"
                for d in display_data:
                    md_table += f"| {d['Nauczyciel']} | {d['E-mail']} | {d['Notatki']} |\n"
                
            st.markdown(md_table, unsafe_allow_html=True)
            
            # --- GENERATOR E-MAILI ---
            st.markdown("#### ✉️ Szybka wiadomość do grupy")
            if f"tpl_{g_idx}" not in st.session_state:
                st.session_state[f"tpl_{g_idx}"] = 0
                
            t_day = g_conf['day'].lower()
            t_start = g_conf['start'].strftime('%H:%M')
            t_end = g_conf['end'].strftime('%H:%M')
            t_branch = g_conf['branch'] if g_conf['branch'] else "wybranej filii"
            
            templates = [
                f"Cześć!\n\nSzukam zastępstwa na {t_day} ({t_start}-{t_end}) w filii {t_branch} dla grupy <TU WPISZ NAZWĘ GRUPY>.\n\nMateriały będą gotowe na miejscu. Ktoś poratuje?\n\nDzięki!",
                f"Hej wszystkim,\n\npotrzebuję pomocy z zastępstwem w najbliższy {t_day}.\nZajęcia: {t_start}-{t_end} w {t_branch} (grupa <TU WPISZ NAZWĘ GRUPY>).\n\nZ góry wielkie dzięki za pomoc!",
                f"Ratunku! Szukam dobrej duszy na zastępstwo.\nKiedy: {t_day}, {t_start}-{t_end}\nGdzie: {t_branch}\nGrupa: <TU WPISZ NAZWĘ GRUPY>\n\nBędę bardzo wdzięczny/a za uratowanie życia!",
                f"Cześć, ma ktoś może wolne okienko w {t_day}?\nSzukam zastępstwa w {t_branch} na godziny {t_start}-{t_end} (grupa <TU WPISZ NAZWĘ GRUPY>).\n\nOdwdzięczę się przy najbliższej okazji! :)",
                f"Hej! Poszukiwane zastępstwo na {t_day} w {t_branch}.\nGodziny: {t_start}-{t_end}\nGrupa: <TU WPISZ NAZWĘ GRUPY>\n\nMateriały zostawię w pełni przygotowane. Pomoże ktoś?",
                f"Cześć! Szukam zastępstwa na {t_day} w {t_branch}. Lekcja trwa od {t_start} do {t_end} dla grupy <TU WPISZ NAZWĘ GRUPY>. Scenariusz będzie czekał. Kto da radę wziąć?\n\nDzięki z góry!"
            ]
            
            current_body = templates[st.session_state[f"tpl_{g_idx}"]]
            
            # Formularz z losowaniem szablonu i przyciskiem Mailto
            col_b1, col_b2 = st.columns([1, 2])
            with col_b1:
                st.button("🎲 Losuj inny tekst", key=f"btn_rand_{g_idx}", on_click=next_tpl, args=(g_idx,))
            with col_b2:
                subject = urllib.parse.quote(f"Zastępstwo - {t_day}")
                body_encoded = urllib.parse.quote(current_body)
                bcc_emails = ";".join(emails)
                mailto_link = f"mailto:?bcc={bcc_emails}&subject={subject}&body={body_encoded}"
                
                st.markdown(
                    f'<a href="{mailto_link}" target="_blank" style="display: inline-block; width: 100%; text-align: center; padding: 0.5em 1em; color: white; background-color: #4CAF50; text-decoration: none; border-radius: 4px; font-weight: bold;">'
                    f'✉️ Wyślij e-mail jednym kliknięciem (otwiera Twoją pocztę)</a>', 
                    unsafe_allow_html=True
                )
            
            st.text_area("Możesz też skopiować tekst ręcznie:", value=current_body, height=180, key=f"text_{g_idx}")
            st.markdown("*(Pamiętaj, by w programie pocztowym podmienić `<TU WPISZ NAZWĘ GRUPY>` i dodać biuro/lidera do DW!)*")
