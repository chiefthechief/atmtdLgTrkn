import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.datavalidation import DataValidation

# --- Configuration ---
EXCEL_FILE = r"{file location for excel file}"
HEADERS = ["Date", "Student/Staff Number", "Issue Reported", "Solution Applied", "State"]

# --- Ensure the directory exists ---
os.makedirs(os.path.dirname(EXCEL_FILE), exist_ok=True)

# --- Load or create workbook (once, before the loop) ---
if os.path.exists(EXCEL_FILE):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)

print("=== IT Support Log Entry Tool ===")
print("Enter your entries below. Type 'quit' at any prompt to exit.\n")

while True:
    print("\n--- New Entry ---")
    
    # Date is automatic
    date_today = datetime.now().strftime("%m/%d/%Y")
    print(f"Date: {date_today}")
    
    # Student number – allow quit
    student_num = input("Student/Staff Number (or 'quit' to exit): ").strip()
    if student_num.lower() == "quit":
        break
    
    # Issue reported
    issue = input("Issue Reported: ").strip()
    if issue.lower() == "quit":
        break
    
    # Solution applied
    solution = input("Solution Applied: ").strip()
    if solution.lower() == "quit":
        break
    
    # State with validation
    state = ""
    while state.lower() not in ["resolved", "pending"]:
        state_input = input("State (resolved/pending) or 'quit': ").strip().lower()
        if state_input == "quit":
            break   # breaks inner loop, but we need to break outer loop too
        if state_input in ["resolved", "pending"]:
            state = state_input
        else:
            print("Please enter either 'resolved' or 'pending'.")
    
    if state == "":
        # User typed quit during state input
        break
    
    # Append the new row
    new_row = [date_today, student_num, issue, solution, state.capitalize()]
    ws.append(new_row)
    
    # --- Add dropdown for the State column (covers all existing data rows) ---
    dv = DataValidation(type="list", formula1='"Resolved,Pending"', allow_blank=True)
    dv.add(f"E2:E{ws.max_row}")
    ws.add_data_validation(dv)
    
    # Save after each entry
    try:
        wb.save(EXCEL_FILE)
        print("✅ Entry saved.")
    except PermissionError:
        print("\n❌ Could not save the file. Please close the Excel file if it's open and try again.")
        # Optionally, you could ask if they want to retry or exit
        break


print("\nExiting. Goodbye!")
