import openpyxl
import pandas as pd
import streamlit as st
from datetime import datetime, time, timedelta

st.set_page_config(page_title="Substitute Teacher Finder", page_icon="📋", layout="wide")
st.title("Substitute Teacher Finder 📋")
st.write("Upload your schedule, pick a day, and set the exact lesson times.")

uploaded_file = st.file_uploader("Upload the Schedule (Excel file)", type=["xlsx"])

# Match the English UI to the Polish sheet names in your Excel file
day_mapping = {
    "Monday": "Poniedziałek",
    "Tuesday": "Wtorek",
    "Wednesday": "Środa",
    "Thursday": "Czwartek",
    "Friday": "Piątek"
}
levels = ["3-5 LAT", "0", "1", "2", "3", "4", "5", "6", "7", "8+", "MASTER"]

col1, col2 = st.columns(2)
with col1:
    selected_day = st.selectbox("Select Day", list(day_mapping.keys()))
with col2:
    selected_level = st.selectbox("Select Level", levels)

st.write("### Lesson Timeframe")
col_t1, col_t2 = st.columns(2)
with col_t1:
    start_time = st.time_input("Lesson Start Time (e.g., 15:50)", value=time(15, 50))
with col_t2:
    end_time = st.time_input("Lesson End Time (e.g., 17:20)", value=time(17, 20))

def is_cell_free(cell):
    """Checks if a cell background is blank (meaning the teacher is free)"""
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
    """Checks if a 15-min Excel block overlaps with the lesson timeframe"""
    dummy_date = datetime(2000, 1, 1)
    block_start = datetime.combine(dummy_date, block_start_time)
    block_end = block_start + timedelta(minutes=15)
    r_start = datetime.combine(dummy_date, req_start)
    r_end = datetime.combine(dummy_date, req_end)
    return block_start < r_end and block_end > r_start

if uploaded_file and st.button("Find Substitutes"):
    try:
        # Load the workbook
        wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        
        # --- 1. Find Qualified Teachers (Sheet 7) ---
        qual_sheet_name = "Lektorzy i poziomy grup na zast"
        uploaded_file.seek(0)
        df_levels = pd.read_excel(uploaded_file, sheet_name=qual_sheet_name, header=None)
        
        qualified_teachers = []
        
        # Find which column holds the requested level
        level_col = None
        for col_idx in df_levels.columns:
            if df_levels[col_idx].head(10).astype(str).str.contains(selected_level, na=False, regex=False).any():
                level_col = col_idx
                break
                
        if level_col is not None:
            for idx, row in df_levels.iterrows():
                if idx < 4: continue # Skip headers
                val = str(row[level_col]).strip().lower()
                if val == "v":
                    teacher_name = str(row[1]).strip() # Column B contains names
                    email = str(row[4]).strip()        # Column E contains emails
                    if teacher_name != "nan":
                        qualified_teachers.append({"Name": teacher_name, "Email": email})
        
        # --- 2. Check Availability (Daily Sheets) ---
        excel_day_name = day_mapping[selected_day]
        day_sheet = wb[excel_day_name]
        
        # Find which 15-minute columns overlap with our lesson time
        overlapping_cols = {}
        for cell in day_sheet[4]: # Row 4 contains the times
            if isinstance(cell.value, time):
                if overlaps(cell.value, start_time, end_time):
                    overlapping_cols[cell.column] = cell.value
        
        available_teachers = []
        
        if not overlapping_cols:
            st.warning("No schedule slots found matching that timeframe.")
        else:
            for teacher in qualified_teachers:
                teacher_row = None
                # Find the teacher in Column A
                for row in range(4, day_sheet.max_row + 1):
                    cell_val = str(day_sheet.cell(row=row, column=1).value).strip()
                    if cell_val != "None" and (teacher["Name"] in cell_val or cell_val in teacher["Name"]):
                        teacher_row = row
                        break
                        
                if teacher_row:
                    # Check all overlapping 15-minute columns
                    is_free = True
                    for col in overlapping_cols.keys():
                        cell = day_sheet.cell(row=teacher_row, column=col)
                        if not is_cell_free(cell):
                            is_free = False
                            break
                            
                    if is_free:
                        available_teachers.append(teacher)
                        
        # --- 3. Show Results ---
        if available_teachers:
            st.success(f"Found {len(available_teachers)} teacher(s) available for the entire lesson!")
            emails = [t["Email"] for t in available_teachers if t["Email"] != "nan" and "@" in t["Email"]]
            email_string = "; ".join(emails)
            
            st.subheader("Copy Emails:")
            st.code(email_string, language="text")
            st.dataframe(pd.DataFrame(available_teachers), hide_index=True)
        else:
            st.warning("No teachers are available and qualified for this exact timeframe.")
            
    except Exception as e:
        st.error(f"Error reading the file: {e}")
