import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime, time, timedelta

st.set_page_config(page_title="Wyszukiwarka Zastępstw", page_icon="📋", layout="wide")
st.title("Wyszukiwarka Zastępstw 📋")
st.write("Wgraj grafik, wybierz dzień, filię i ustal dokładne godziny zajęć.")

uploaded_file = st.file_uploader("Wgraj grafik (plik Excel)", type=["xlsx"])

days = ["Poniedziałek", "Wtorek", "Środa", "Czwartek", "Piątek"]
levels = ["3-5 LAT", "0", "1", "2", "3", "4", "5", "6", "7", "8+", "MASTER"]

# Mapowanie dni na nowe skrócone nazwy arkuszy w pliku 2026/27
day_mapping = {
    "Poniedziałek": "PON",
    "Wtorek": "WT",
    "Środa": "ŚR",
    "Czwartek": "CZW",
    "Piątek": "PT"
}

# Mapowanie opcji wyboru na skróty z pliku Excel
branches = {
    "Dowolna (szukaj wszystkich)": None,
    "Ursus 1 (U1)": "U1",
    "Ursus 2 (U2)": "U2",
    "Komorów (K)": "K",
    "Michałowice (M)": "M",
    "Nowa Wieś (NW)": "NW",
    "Pruszków (P)": "P"
}

col1, col2, col3 = st.columns(3)
with col1:
    selected_day = st.selectbox("Wybierz dzień", days)
with col2:
    selected_level = st.selectbox("Wybierz poziom", levels)
with col3:
    selected_branch = st.selectbox("Wybierz filię (opcjonalnie)", list(branches.keys()))

st.write("### Czas trwania zajęć")
col_t1, col_t2 = st.columns(2)
with col_t1:
    start_time = st.time_input("Czas rozpoczęcia (np. 15:50)", value=time(15, 50))
with col_t2:
    end_time = st.time_input("Czas zakończenia (np. 17:20)", value=time(17, 20))

def is_cell_free(cell):
    """Sprawdza, czy tło komórki jest puste (brak koloru = wolny)"""
    if not cell.fill or cell.fill.fill_type is None:
        return True
    color = getattr(cell.fill.start_color, "index", None)
    if color in [0, "00000000", "FFFFFFFF", "None", None]:
        return True
    rgb = getattr(cell.fill.start_color, "rgb", None)
    if rgb in ["00000000", "FFFFFFFF", None]:
        return True
    return False

def overlaps(block_start_time, req_start, req_end):
    """Sprawdza, czy 15-minutowy blok nachodzi na godziny lekcji"""
    dummy_date = datetime(2000, 1, 1)
    block_start = datetime.combine(dummy_date, block_start_time)
    block_end = block_start + timedelta(minutes=15)
    r_start = datetime.combine(dummy_date, req_start)
    r_end = datetime.combine(dummy_date, req_end)
    return block_start < r_end and block_end > r_start

if uploaded_file and st.button("Znajdź Zastępstwo"):
    try:
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # --- 1. Sprawdzenie kwalifikacji i filii (Arkusz Lektorzy i poziomy...) ---
        qual_sheet_name = "Lektorzy i poziomy grup na zast"
        uploaded_file.seek(0)
        df_levels = pd.read_excel(uploaded_file, sheet_name=qual_sheet_name, header=None)
        
        qualified_teachers = []
        
        level_col = None
        for col_idx in df_levels.columns:
            if df_levels[col_idx].head(10).astype(str).str.contains(selected_level, na=False, regex=False).any():
                level_col = col_idx
                break
                
        if level_col is not None:
            for idx, row in df_levels.iterrows():
                if idx < 4: continue 
                val = str(row[level_col]).strip().lower()
                if val == "v":
                    teacher_name = str(row[1]).strip() 
                    branch_val = str(row[2]).strip()   
                    email = str(row[4]).strip()        
                    
                    if teacher_name != "nan":
                        is_at_branch = False
                        selected_branch_code = branches[selected_branch]
                        
                        if selected_branch_code is not None:
                            branch_upper = branch_val.upper()
                            if "KUM" in branch_upper and "CO" in branch_upper:
                                is_at_branch = True
                            else:
                                # Obsługa zarówno "NW+P", jak i "NW&P"
                                teacher_branches = [b.strip() for b in branch_upper.replace(" ", "").replace("&", "+").split("+")]
                                if selected_branch_code in teacher_branches:
                                    is_at_branch = True
                        
                        notes = "Uczę w tej filii :)" if is_at_branch else ""
                        
                        qualified_teachers.append({
                            "Nauczyciel": teacher_name,
                            "E-mail": email,
                            "Filie wg tabeli": branch_val,
                            "Notatki": notes
                        })
        
        # --- 2. Sprawdzenie dostępności w grafiku (Arkusze dni) ---
        excel_day_name = day_mapping[selected_day]
        day_sheet = wb[excel_day_name]
        
        overlapping_cols = {}
        for cell in day_sheet[4]: 
            if isinstance(cell.value, time):
                if overlaps(cell.value, start_time, end_time):
                    overlapping_cols[cell.column] = cell.value
        
        available_teachers = []
        
        if not overlapping_cols:
            st.warning("Nie znaleziono w pliku kolumn z godzinami pasującymi do tego przedziału czasowego.")
        else:
            for teacher in qualified_teachers:
                teacher_row = None
                for row in range(4, day_sheet.max_row + 1):
                    cell_val = str(day_sheet.cell(row=row, column=1).value).strip()
                    if cell_val != "None" and (teacher["Nauczyciel"] in cell_val or cell_val in teacher["Nauczyciel"]):
                        teacher_row = row
                        break
                        
                if teacher_row:
                    is_free = True
                    for col in overlapping_cols.keys():
                        cell = day_sheet.cell(row=teacher_row, column=col)
                        if not is_cell_free(cell):
                            is_free = False
                            break
                            
                    if is_free:
                        available_teachers.append(teacher)
                        
        # --- 3. Wyświetlenie wyników ---
        if available_teachers:
            st.success(f"Znaleziono dostępnych nauczycieli: {len(available_teachers)}!")
            
            emails = [t["E-mail"] for t in available_teachers if t["E-mail"] != "nan" and "@" in t["E-mail"]]
            email_string = "; ".join(emails)
            
            st.subheader("Skopiuj adresy e-mail (do wklejenia w Outlook/Gmail):")
            st.code(email_string, language="text")
            
            st.subheader("Szczegóły:")
            st.dataframe(pd.DataFrame(available_teachers), hide_index=True)
        else:
            st.warning("Brak nauczycieli, którzy mają wolny czas i odpowiednie kwalifikacje w podanym przedziale czasowym.")
            
    except Exception as e:
        st.error(f"Wystąpił błąd podczas odczytu pliku: {e}")
