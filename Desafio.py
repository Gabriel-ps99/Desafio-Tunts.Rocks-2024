import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time

# Authentication settings
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

# Get the credentials file path from the environment variable
cred_path = os.getenv("CREDENTIALS_PATH")

# Check if the environment variable is set
if not cred_path:
    print("Error: The environment variable CREDENTIALS_PATH is not set.")
    exit()

creds = ServiceAccountCredentials.from_json_keyfile_name(cred_path, scope)

# Exception handling for authentication
try:
    client = gspread.authorize(creds)
except gspread.exceptions.SpreadsheetNotFound as e:
    print(f"Authentication error: {e}")
    exit()

# Open the spreadsheet by ID
spreadsheet_id = "1XlGYFv5PqnEXVGqvy0hxbeLQlir35PKQGzgO7nbBS4s"
spreadsheet = client.open_by_key(spreadsheet_id)

# Select the sheet/tab of the spreadsheet
worksheet = spreadsheet.sheet1  # or use worksheet = spreadsheet.get_worksheet(0) to select by index

# Get all values from the spreadsheet
all_values = worksheet.get_all_values()

# Iterate over the rows of the spreadsheet (ignoring the first three rows)
total_aulas = 60  # replace with the total number of classes in the semester
for row_index, row in enumerate(all_values[3:], start=4):
    # Extract information from the row
    matricula, aluno_nome, faltas, p1, p2, p3, situacao, naf = row

    # Check if the columns for grades and absences contain numeric values
    if faltas.replace('.', '', 1).isdigit() and p1.replace('.', '', 1).isdigit() and p2.replace('.', '', 1).isdigit() and p3.replace('.', '', 1).isdigit():
        # Convert grades and absences to float
        faltas, p1, p2, p3 = float(faltas), float(p1), float(p2), float(p3)

        # Calculate the average correctly
        media = round((p1 + p2 + p3) / 3, 1)

        # Check the student's situation
        if faltas > 0.25 * total_aulas:
            situacao = "Reprovado por Falta"
            naf = 0  # If failed due to absence, naf is zero
        elif media < 5:
            situacao = "Reprovado por Nota"
            naf = 0  # If failed due to grades, naf is zero
        elif 50 <= media < 70:
            situacao = "Exame Final"
            # Calculate the Final Approval Grade (naf)
            naf = max(2 * 5 - media, 0)
        else:
            situacao = "Aprovado"
            naf = 0

        # Add log lines
        log = f"Student: {aluno_nome}, Average: {media / 10:.2f}, Situation: {situacao}, Absences: {faltas}, NAF: {10 - media / 10:.2f}"
        print(log)

        # Update the spreadsheet with the new information
        worksheet.update_cell(row_index, 7, situacao)  # Update the Situation column
        if situacao != "Exame Final":
            worksheet.update_cell(row_index, 8, "0.00")
        else:
            worksheet.update_cell(row_index, 8, f"{10 - media / 10:.2f}")  # Update the Final Grade column
    
        # Add a small delay (1 second) between updates
        time.sleep(1)
    else:
        print(f"Ignoring header or error in row {row_index}: One of the grade or absence columns is empty or not numeric.")

print("Processing completed.")
