import os
import pandas as pd
from openpyxl import load_workbook
import win32com.client as win32

# Folder path with Excel files
excel_folder = "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/Phyton_task/IE3_DOWNLOAD"

# List of CSV files with passwords
password_files = [
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_1000_Paula.csv",
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_2500_Paula.csv",
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_5000_Paula.csv"
]

# Read passwords from CSV files
passwords = []
for password_file in password_files:
    df = pd.read_csv(password_file)
    passwords.extend(df['Passwords'].tolist())  # Change 'Passwords' to the correct column name

print("Passwords obtained:", passwords)  # Verify that passwords are read correctly

# Encrypt Excel files in the folder
for root, dirs, files in os.walk(excel_folder):
    for file in files:
        if file.endswith(".xlsx"):
            excel_file_path = os.path.join(root, file)
            try:
                password = passwords.pop(0)  # Get the next password from the list
                wb = load_workbook(excel_file_path)
                ws = wb.active
                ws.protection.set_password(password)
                wb.save(excel_file_path)

                print(f"File encrypted: {excel_file_path}")
                
            except Exception as e:
                print(f"Error encrypting {excel_file_path}: {str(e)}")

print("Encryption process completed.")







