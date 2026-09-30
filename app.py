import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime, time, timedelta
import os

st.set_page_config(page_title="Wyszukiwarka Zastępstw", page_icon="📋", layout="wide")
st.title("Wyszukiwarka Zastępstw 📋")

# --- OBSŁUGA PLIKU (DOMYŚLNY LUB WGRANY RĘCZNIE) ---
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

# --- USTAWIENIA NA GŁÓWNYM EKRANIE ---
st.markdown("---")
st.write("### Opcje wyszukiwania")
col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    exclude_me = st.text_input("Wyklucz mnie (Twoje Imię i Nazwisko)", placeholder="np. Jan Kowalski").strip().lower()
with col_opt2:
    st.write("") 
    is_multiple = st.checkbox("Szukam zastępstw dla więcej niż 1 grupy", value=False)
    if is_multiple:
        num_groups = st.number_input("Ile grup?", min_value=2, max_value=5, value=2)
    else:
        num_groups = 1

st.markdown("---")

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

# --- KONFIGURACJA GRUP ---
groups_config = []
st.write("### Parametry Grupy 1")
col1, col2, col3 = st.columns(3)
with col1: default_day = st.selectbox("Dzień", days, key="day_0")
with col2: default_level = st.selectbox("Poziom", levels, key="level_0")
with col3: default_branch = st.selectbox("Która filia?", list(branches.keys()), key="branch_0")

col_t1, col_t2 = st.columns(2)
with col_t1: 
    default_start = st.time_input("Czas rozpoczęcia", value=time(15, 50), key="start_0")

# Automatyczne dodawanie 90 minut
start_dt = datetime.combine(datetime.today(), default_start)
auto_end = (start_dt + timedelta(minutes=90)).time()

with col_t2: 
    default_end = st.time_input("Czas zakończenia", value=auto_end, key="end_0")

groups_config.append({
    "day": default_day, "level": default_level, "branch": default_branch, 
    "start": default_start, "end": default_end
})

if is_multiple:
    for i in range(1, int(num_groups)):
        st.markdown("---")
        st.write(f"### Parametry Grupy {i+1}")
        c1, c2, c3 = st.columns(3)
        with c1: g_day = st.selectbox("Dzień", days, index=days.index(default_day), key=f"day_{i}")
        with c2: g_level = st.selectbox("Poziom", levels, key=f"level_{i}")
        with c3: g_branch = st.selectbox("Która filia?", list(branches.keys()), index=list(branches.keys()).index(default_branch), key=f"branch_{i}")
        
        ct1, ct2 = st.columns(2)
        with ct1: 
            g_start = st.time_input("Czas rozpoczęcia", value=default_start, key=f"start_{i}")
            
        g_start_dt = datetime.combine(datetime.today(), g_start)
        g_auto_end = (g_start_dt + timedelta(minutes=90)).time()
            
        with ct2: 
            g_end = st.time_input("Czas zakończenia", value=g_auto_end, key=f"end_{i}")
        
        groups_config.append({
            "day": g_day, "level": g_level, "branch": g_branch, 
            "start": g_start, "end": g_end
        })

# --- FUNKCJE POMOCNICZE ---
def is_cell_free(cell):
    if not cell.fill or cell.fill.fill_type is None: return True
    color = getattr(cell.fill.start_color, "index", None)
    if color in [0, "00000000", "FFFFFFFF", "None", None]: return True
    rgb = getattr(cell.fill.start_color, "rgb", None)
    if rgb in ["00000000", "FFFFFFFF", None]: return True
    return False

def overlaps(block_start_time, req_start, req_end):
    dummy_date = datetime(2000, 1, 1)
    block_start = datetime.combine(dummy_date, block_start_time)
    block_end = block_start + timedelta(minutes=15)
    r_start = datetime.combine(dummy_date, req_start)
    r_end = datetime.combine(dummy_date, req_end)
    return block_start < r_end and block_end > r_start

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
            if col_dt < req_dt:
                t_cell = day_sheet.cell(row=teacher_row, column=cell.column)
                if not is_cell_free(t_cell):
                    last_busy_end = col_dt + timedelta(minutes=15)
                    
    if last_busy_end and last_busy_end <= req_dt:
        diff_mins = int((req_dt - last_busy_end).total_seconds() / 60)
        if 0 <= diff_mins <= 60:
            od = loc_grammar[t_branch]
            do = loc_grammar[target_branch_code]
            return f"Kończę zajęcia o {last_busy_end.strftime('%H:%M')}. Mam tylko {diff_mins} min na dojazd z {od} do {do}."
    return ""

st.markdown("---")

# --- ZARZĄDZANIE PAMIĘCIĄ SESJI (Żeby filtry nie resetowały wyników) ---
if "search_results" not in st.session_state:
    st.session_state.search_results = None
    st.session_state.teacher_counts = {}

# --- GŁÓWNY PROCES ---
if st.button("Znajdź Zastępstwo 🚀", use_container_width=True):
    if active_file is None:
        st.error("⚠️ Proszę najpierw wgrać plik z grafikiem!")
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
                    if df_levels[col_idx].head(10).astype(str).str.contains(g_conf["level"], na=False, regex=False).any():
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
                            email = str(row[4]).strip()
                            
                            if t_name != "nan":
                                q_teachers.append({"name": t_name, "email": email, "branch": b_val})
                                
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
                                    "Filie": t["branch"],
                                    "_is_at_branch": is_at_branch,
                                    "_commute_warn": commute_warn
                                })
                                
                all_group_results.append(a_teachers)
                
            # Zliczanie ilości grup dla lektora
            counts = {}
            for res in all_group_results:
                for t in res:
                    counts[t["Nauczyciel"]] = counts.get(t["Nauczyciel"], 0) + 1
                    
            # Zapis do pamięci sesji
            st.session_state.search_results = all_group_results
            st.session_state.teacher_counts = counts
            
        except Exception as e:
            st.error(f"Wystąpił błąd: {e}")

# --- WYSWIETLANIE I FILTROWANIE WYNIKÓW ---
if st.session_state.search_results is not None:
    st.write("### ⚙️ Filtry wyników")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        if is_multiple:
            show_only_all = st.toggle("Pokaż tylko lektorów, którzy mogą wziąć WSZYSTKIE zastępstwa")
        else:
            show_only_all = False
    with col_f2:
        exclude_bad_commute = st.toggle("Wyklucz lektorów, którzy mają mało czasu na dojazd")

    custom_css = """
    <style>
    .styled-table { border-collapse: collapse; margin: 15px 0; width: 100%; font-family: inherit; font-size: 0.9em; }
    .styled-table thead tr { background-color: rgba(128, 128, 128, 0.15); text-align: left; }
    .styled-table th, .styled-table td { padding: 10px 15px; border-bottom: 1px solid rgba(128,128,128,0.2); }
    </style>
    """

    for g_idx, (g_conf, res) in enumerate(zip(groups_config, st.session_state.search_results)):
        st.markdown(f"## Wyniki dla Grupy {g_idx+1}")
        st.write(f"**{g_conf['day']} | {g_conf['start'].strftime('%H:%M')} - {g_conf['end'].strftime('%H:%M')} | Poziom: {g_conf['level']} | Filia: {g_conf['branch']}**")
        
        display_data = []
        for t in res:
            count = st.session_state.teacher_counts[t["Nauczyciel"]]
            
            if show_only_all and count < int(num_groups):
                continue
            if exclude_bad_commute and t["_commute_warn"] != "":
                continue
                
            notes_html = []
            
            if t["_is_at_branch"]:
                notes_html.append('<span title="Uczę w tej samej filii :)" style="cursor:help; font-size:1.4em; margin-right:5px;">🏫</span>')
            
            if t["_commute_warn"]:
                notes_html.append(f'<span title="{t["_commute_warn"]}" style="cursor:help; font-size:1.4em; margin-right:5px;">⚠️</span>')
                
            if is_multiple:
                if count == int(num_groups):
                    notes_html.append(f'<span title="Mogę wziąć WSZYSTKIE {count} zastępstwa!" style="cursor:help; font-size:1.4em; margin-right:5px;">🔥</span>')
                elif count > 1:
                    notes_html.append(f'<span title="Mogę wziąć {count} zastępstwa." style="cursor:help; font-size:1.4em; margin-right:5px;">⭐</span>')
                    
            notes_str = "".join(notes_html) if notes_html else "-"
            
            display_data.append({
                "Nauczyciel": t["Nauczyciel"],
                "E-mail": t["E-mail"],
                "Notatki": notes_str
            })
            
        if not display_data:
            st.warning("Brak nauczycieli spełniających wybrane kryteria i filtry.")
        else:
            emails = [d["E-mail"] for d in display_data if d["E-mail"] != "nan" and "@" in d["E-mail"]]
            st.code("; ".join(emails), language="text")
            
            df_display = pd.DataFrame(display_data)
            html_table = df_display.to_html(escape=False, index=False)
            html_table = html_table.replace('<table border="1" class="dataframe">', '<table class="styled-table">')
            st.markdown(custom_css + html_table, unsafe_allow_html=True)
