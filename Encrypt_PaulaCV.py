import aspose.words as aw
import os

# Input file path
input_file_path = "C:\\Users\\Administrator\\OneDrive - School of Automation\\Blue Prism Advance SOA\\PythonEncryptionProject\\Phyton_Encryption_Files\\Paula_cv.docx"

# Output file path with password
output_file_path = "C:\\Users\\Administrator\\OneDrive - School of Automation\\Blue Prism Advance SOA\\Phyton_task\\Phyton_Encryption_Files\\document-password-protected.docx"

# Load the document
doc = aw.Document(input_file_path)

# Create document options
options = aw.saving.OoxmlSaveOptions(aw.SaveFormat.DOCX)

# Set password
options.password = "Paula10"

# Save the updated document
doc.save(output_file_path, options)

print(f"Document with password saved in: {output_file_path}")




