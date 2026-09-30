import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Substitute Teacher Finder", page_icon="📋")
st.title("Substitute Teacher Finder 📋")
st.write("Upload your schedule to find available and qualified teachers.")

uploaded_file = st.file_uploader(
    "Upload the Schedule (Excel file)", type=["xlsx"]
)

days_of_week = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
times = [
    "08:00",
    "08:15",
    "08:30",
    "08:45",
    "09:00",
    "09:15",
    "09:30",
    "09:45",
    "10:00",
    "10:15",
    "10:30",
    "10:45",
    "11:00",
    "11:15",
    "11:30",
    "11:45",
    "12:00",
    "12:15",
    "12:30",
    "12:45",
    "13:00",
    "13:15",
    "13:30",
    "13:45",
    "14:00",
    "14:15",
    "14:30",
    "14:45",
    "15:00",
    "15:15",
    "15:30",
    "15:45",
    "16:00",
    "16:15",
    "16:30",
    "16:45",
    "17:00",
    "17:15",
    "17:30",
    "17:45",
    "18:00",
    "18:15",
    "18:30",
    "18:45",
    "19:00",
    "19:15",
    "19:30",
    "19:45",
    "20:00",
    "20:15",
    "20:30",
    "20:45",
    "21:00",
]
levels = ["3-5 LAT", "0", "1", "2", "3", "4", "5", "6", "7"]

col1, col2, col3 = st.columns(3)
with col1:
  selected_day = st.selectbox("Select Day", days_of_week)
with col2:
  selected_time = st.selectbox("Select Time", times)
with col3:
  selected_level = st.selectbox("Select Level", levels)


def is_cell_free(cell):
  """Checks if a cell is white, transparent, or uncolored."""
  if not cell.fill or cell.fill.fill_type is None:
    return True
  color = getattr(cell.fill.start_color, "index", None)
  if color in [0, "00000000", "FFFFFFFF", "None", None]:
    return True
  rgb = getattr(cell.fill.start_color, "rgb", None)
  if rgb in ["00000000", "FFFFFFFF", None]:
    return True
  return False


if uploaded_file and st.button("Find Substitutes"):
  try:
    # Read Sheet 6 (Qualifications)
    uploaded_file.seek(0)
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    sheet_names = wb.sheetnames
    levels_sheet_name = sheet_names[5]

    uploaded_file.seek(0)
    df_levels = pd.read_excel(
        uploaded_file, sheet_name=levels_sheet_name, header=None
    )

    qualified_teachers = []
    level_col = None
    for col in df_levels.columns:
      if (
          df_levels[col]
          .astype(str)
          .str.contains(selected_level, na=False)
          .any()
      ):
        level_col = col
        break

    if level_col is not None:
      for _, row in df_levels.iterrows():
        val = str(row[level_col]).strip().lower()
        if val == "v":
          teacher_name = str(row[0]).strip()
          email = str(row[3]).strip()
          qualified_teachers.append({"Name": teacher_name, "Email": email})

    # Read Daily Availability Sheets (Sheets 1 to 5)
    day_index = days_of_week.index(selected_day)
    day_sheet = wb[sheet_names[day_index]]

    time_col_idx = None
    for cell in day_sheet[4]:
      if cell.value is not None and selected_time in str(cell.value):
        time_col_idx = cell.column
        break

    available_teachers = []
    if time_col_idx:
      for teacher in qualified_teachers:
        teacher_row = None
        for row in range(5, day_sheet.max_row + 1):
          cell_val = day_sheet.cell(row=row, column=1).value
          if cell_val is not None and teacher["Name"] in str(cell_val):
            teacher_row = row
            break

        if teacher_row:
          cell = day_sheet.cell(row=teacher_row, column=time_col_idx)
          if is_cell_free(cell):
            available_teachers.append(teacher)

    # Show results
    if available_teachers:
      st.success(
          f"Found {len(available_teachers)} teacher(s) for {selected_level} on"
          f" {selected_day} at {selected_time}!"
      )

      emails = [
          t["Email"]
          for t in available_teachers
          if t["Email"] != "nan" and "@" in t["Email"]
      ]
      email_string = "; ".join(emails)

      st.subheader("Copy Emails:")
      st.code(email_string, language="text")

      st.subheader("Teacher Details:")
      st.dataframe(pd.DataFrame(available_teachers), hide_index=True)
    else:
      st.warning("No teachers are available and qualified for this slot.")

  except Exception as e:
    st.error(
        "There was an error reading the file. Make sure your Excel columns"
        f" match the standard layout. Details: {e}"
    )
