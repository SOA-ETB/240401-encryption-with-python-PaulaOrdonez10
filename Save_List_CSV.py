import os
import pandas as pd
from openpyxl import load_workbook

# Nueva ruta con archivos de Excel encriptados
excel_folder = "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PythonEncryptionProject/Phyton_Encryption_Files/IE3_DOWNLOAD"

# List of CSV with passwords
password_files = [
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_1000_Paula.csv",
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_2500_Paula.csv",
    "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PasswordGeneratorProject/passwords_5000_Paula.csv"
]

# Leer contraseñas de los archivos CSV
passwords = []
for password_file in password_files:
    df = pd.read_csv(password_file)
    passwords.extend(df['Passwords'].tolist()) 

print("Passwords obtained:", passwords)  # Verificar que las contraseñas se lean correctamente

# Crear un diccionario para almacenar los nombres de archivo de los archivos de Excel encriptados y sus contraseñas
encrypted_files = {}

# Encriptar archivos de Excel en la carpeta y almacenar nombres de archivo y contraseñas en el diccionario
for root, dirs, files in os.walk(excel_folder):
    for file in files:
        if file.endswith(".xlsx"):
            excel_file_path = os.path.join(root, file)
            try:
                password = passwords.pop(0)  # Obtener la próxima contraseña
                wb = load_workbook(excel_file_path)
                ws = wb.active
                ws.protection.set_password(password)
                wb.save(excel_file_path)
                encrypted_files[excel_file_path] = password
                print(f"File encrypted: {excel_file_path}")
            except Exception as e:
                print(f"Error encrypting {excel_file_path}: {str(e)}")

# Guardar el diccionario en un archivo CSV
output_csv = "C:/Users/Administrator/OneDrive - School of Automation/Blue Prism Advance SOA/PythonEncryptionProject/Phyton_Encryption_Files/encrypted_files.csv"
df_output = pd.DataFrame(encrypted_files.items(), columns=['Filename', 'Password'])
df_output.to_csv(output_csv, index=False)
print(f"Encrypted files list saved to {output_csv}")


