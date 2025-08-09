import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import io  # Used to handle the file in memory
import os  # Used to handle file names

# =====================================================================================
#  OUR RESCHEDULING LOGIC (ADAPTED INSIDE A FUNCTION)
# =====================================================================================
def process_spreadsheet(excel_file, holiday_day):
    """
    This function contains our main algorithm.
    It receives the uploaded file and the holiday day, and returns
    the modified file and a list of logs.
    """
    logs = []  # List to store log messages

    try:
        # We use the in-memory file directly
        workbook = openpyxl.load_workbook(excel_file)
        sheet_name = '01. Calendario SCL Abarrotes'
        sheet = workbook[sheet_name]
        logs.append(f"Spreadsheet '{excel_file.name}' loaded successfully.")
    except Exception as e:
        logs.append(f"ERROR: Could not read the spreadsheet. Please check if it is the correct file. Details: {e}")
        return None, logs

    delivery_columns = ['AI', 'AJ', 'AK', 'AL', 'AM', 'AN']
    observations_column = 'CT'
    weekday_map = {'L': 1, 'M': 2, 'W': 3, 'J': 4, 'V': 5, 'S': 6, 'D': 7}

    holiday_col_letter = None
    for col_letter in delivery_columns:
        day_in_sheet = sheet[f'{col_letter}3'].value
        if day_in_sheet == holiday_day:
            holiday_col_letter = col_letter
            break
    
    if not holiday_col_letter:
        logs.append(f"ERROR: The day {holiday_day} was not found in row 3 of the Delivery columns.")
        return None, logs
    
    logs.append(f"Holiday identified in the Delivery column: {holiday_col_letter}")

    holiday_col_index = column_index_from_string(holiday_col_letter)
    if holiday_col_index == column_index_from_string(delivery_columns[0]):
         logs.append("Warning: The holiday is the first day of the period. It cannot be anticipated.")
         return None, logs

    previous_col_index = holiday_col_index - 1
    previous_col_letter = get_column_letter(previous_col_index)
    
    tasks_moved = 0
    
    for row_index in range(8, sheet.max_row + 1):
        task_cell = sheet[f'{holiday_col_letter}{row_index}']
        if isinstance(task_cell.value, (int, float)) and 1 <= task_cell.value <= 6:
            destination_cell = sheet[f'{previous_col_letter}{row_index}']
            weekday_initial = sheet[f'{previous_col_letter}6'].value.upper()
            new_weekday_number = weekday_map.get(weekday_initial)
            
            if new_weekday_number:
                destination_cell.value = new_weekday_number
                task_cell.value = None
                log_message = f"Delivery rescheduled (with substitution) from day {holiday_day} to column {previous_col_letter}."
                sheet[f'{observations_column}{row_index}'].value = log_message
                tasks_moved += 1
    
    logs.append(f"Rescheduling completed. {tasks_moved} tasks were moved.")
    
    return workbook, logs

# =====================================================================================
#  STREAMLIT WEB INTERFACE
# =====================================================================================

st.title("🤖 Automatic Holiday Rescheduler")

st.write("""
This tool automates the rescheduling of deliveries in logistics spreadsheets.
Just upload your spreadsheet, enter the holiday day, and click 'Reschedule'.
""")

uploaded_file = st.file_uploader(
    "1. Choose your scheduling spreadsheet (.xlsx)",
    type=['xlsx']
)

holiday_day = st.number_input(
    "2. Enter the day of the month that is a holiday (e.g., 20)",
    min_value=1, 
    max_value=31, 
    step=1,
    value=20  # Default value to facilitate testing
)

if st.button("Reschedule Spreadsheet"):
    if uploaded_file is not None:
        with st.spinner('Please wait... Rescheduling tasks...'):
            modified_workbook, logs = process_spreadsheet(uploaded_file, int(holiday_day))

        st.subheader("Operation Report:")
        for log in logs:
            st.info(log)
        
        if modified_workbook:
            output = io.BytesIO()
            modified_workbook.save(output)
            output.seek(0)
            
            st.success("Your spreadsheet has been successfully rescheduled!")
            
            # --- FILE NAME LOGIC CHANGED HERE ---
            # 1. Get the original file name (e.g., 'spreadsheet.xlsx')
            original_filename = uploaded_file.name
            # 2. Split the base name and extension (e.g., 'spreadsheet', '.xlsx')
            base_name, extension = os.path.splitext(original_filename)
            # 3. Create the new name with suffix
            new_filename = f"{base_name}_rescheduled{extension}"
            
            st.download_button(
                label="Click here to download the rescheduled spreadsheet",
                data=output,
                # 4. Use the new file name on the download button
                file_name=new_filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    else:
        st.error("Please upload a spreadsheet before rescheduling.")
