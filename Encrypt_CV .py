import aspose.words as aw

# Ruta del archivo de entrada
input_file_path = "C:\\Users\\Administrator\\OneDrive - School of Automation\\Blue Prism Advance SOA\\Phyton_task\\python-docx\\cv.docx"

# Ruta del archivo de salida con contraseña
output_file_path = "C:\\Users\\Administrator\\OneDrive - School of Automation\\Blue Prism Advance SOA\\Phyton_task\\python-docx\\document-password-protected.docx"

# load document
doc = aw.Document(input_file_path)

# create document options
options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)

# set password
options.password = "Paula10"

# save updated document
doc.save(output_file_path, options)

print(f"Document with password saved in: {output_file_path}")

