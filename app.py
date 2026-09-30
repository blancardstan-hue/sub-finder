import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime, time, timedelta
import os

st.set_page_config(page_title="Wyszukiwarka Zastępstw", page_icon="📋", layout="wide")
st.title("Wyszukiwarka Zastępstw 📋")

# --- SPRAWDZANIE PLIKU NA SERWERZE (Brak konieczności wgrywania!) ---
default_file = "Zastępstwa 2026_27 - KUM&Co.xlsx"
if os.path.exists(default_file):
    st.success(f"✅ Używam pliku z serwera: **{default_file}**. Nie musisz niczego wgrywać!")
    uploaded_file = default_file
else:
    uploaded_file = st.file_uploader("Wgraj grafik (plik Excel)", type=["xlsx"])

# --- USTAWIENIA NA GŁÓWNYM EKRANIE ---
st.markdown("---")
st.write("### Opcje wyszukiwania")
col_opt1, col_opt2 = st.columns(2)
with col_opt1:
    exclude_me = st.text_input("Wyklucz mnie (Twoje Imię i Nazwisko)", placeholder="np. Jan Kowalski").strip().lower()
with col_opt2:
    st.write("") # Drobne wyrównanie
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
    "Dowolna (szukaj wszystkich)": None,
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
with col3: default_branch = st.selectbox("Filia docelowa", list(branches.keys()), key="branch_0")

col_t1, col_t2 = st.columns(2)
with col_t1: default_start = st.time_input("Czas rozpoczęcia", value=time(15, 50), key="start_0")
with col_t2: default_end = st.time_input("Czas zakończenia", value=time(17, 20), key="end_0")

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
        with c3: g_branch = st.selectbox("Filia docelowa", list(branches.keys()), index=list(branches.keys()).index(default_branch), key=f"branch_{i}")
        
        ct1, ct2 = st.columns(2)
        with ct1: g_start = st.time_input("Czas rozpoczęcia", value=default_start, key=f"start_{i}")
        with ct2: g_end = st.time_input("Czas zakończenia", value=default_end, key=f"end_{i}")
        
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
            return f"⚠️ Kończę zajęcia o {last_busy_end.strftime('%H:%M')}. Mam tylko {diff_mins} min, żeby przejechać z {od} do {do}, mogę mieć problem, żeby zdążyć."
    return ""

st.markdown("---")
# --- GŁÓWNY PROCES ---
if st.button("Znajdź Zastępstwo 🚀", use_container_width=True):
    if uploaded_file is None:
        st.error("⚠️ Proszę najpierw wgrać plik z grafikiem!")
    else:
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            if not isinstance(uploaded_file, str): 
                uploaded_file.seek(0)
                
            df_levels = pd.read_excel(uploaded_file, sheet_name="Lektorzy i poziomy grup na zast", header=None)
            
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
                
            teacher_counts = {}
            for res in all_group_results:
                for t in res:
                    teacher_counts[t["Nauczyciel"]] = teacher_counts.get(t["Nauczyciel"], 0) + 1
                    
            for g_idx, (g_conf, res) in enumerate(zip(groups_config, all_group_results)):
                st.markdown(f"## Wyniki dla Grupy {g_idx+1}")
                st.write(f"**{g_conf['day']} | {g_conf['start'].strftime('%H:%M')} - {g_conf['end'].strftime('%H:%M')} | Poziom: {g_conf['level']} | Filia: {g_conf['branch']}**")
                
                if not res:
                    st.warning("Brak nauczycieli spełniających kryteria dla tej grupy.")
                    continue
                    
                display_data = []
                for t in res:
                    notes = []
                    if t["_is_at_branch"]:
                        notes.append("Uczę w tej filii :)")
                    if t["_commute_warn"]:
                        notes.append(t["_commute_warn"])
                        
                    count = teacher_counts[t["Nauczyciel"]]
                    if is_multiple:
                        if count == int(num_groups):
                            notes.append(f"🔥 Mogę wziąć WSZYSTKIE {int(num_groups)} zastępstwa!")
                        elif count > 1:
                            notes.append(f"Mogę wziąć {count} zastępstwa.")
                            
                    display_data.append({
                        "Nauczyciel": t["Nauczyciel"],
                        "E-mail": t["E-mail"],
                        "Notatki": " | ".join(notes)
                    })
                    
                emails = [t["E-mail"] for t in display_data if t["E-mail"] != "nan" and "@" in t["E-mail"]]
                email_string = "; ".join(emails)
                st.code(email_string, language="text")
                st.dataframe(pd.DataFrame(display_data), hide_index=True)
                
        except Exception as e:
            st.error(f"Wystąpił błąd: {e}")
