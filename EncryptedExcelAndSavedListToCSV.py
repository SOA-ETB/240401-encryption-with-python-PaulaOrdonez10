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
encrypted_files = []  # List to store filenames and passwords

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
                encrypted_files.append((excel_file_path, password))
                print(f"File encrypted: {excel_file_path}")
                
                # Use pywin32 to set Excel file password protection
                excel = win32.Dispatch('Excel.Application')
                excel.Visible = True  # Open Excel visibly
                wb = excel.Workbooks.Open(excel_file_path, False, True, None, password)
                wb.Close(True)  # Save and close Excel file
                
            except Exception as e:
                print(f"Error encrypting {excel_file_path}: {str(e)}")

# Save encrypted files list to CSV in the desired folder
output_csv = "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/Phyton_task/python-docx/encrypted_files.csv"
df_output = pd.DataFrame(encrypted_files, columns=['Filename', 'Password'])
df_output.to_csv(output_csv, index=False)
print(f"Encrypted files list saved to {output_csv}")








