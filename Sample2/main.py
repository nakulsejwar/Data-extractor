import re
import os
import pandas as pd

# Define regular expressions for email and phone number extraction
email_regex = r'\b[A-Za-z0-9._]+@[A-Za-z0-9]+\.[A-Z|a-z]{2,}\b'
phone_regex = r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b'

# Function to extract information from a single CV file
def extract_info(file_path):
    with open(file_path, 'r', encoding='latin-1') as file:
        text = file.read()
        emails = re.findall(email_regex, text)
        phones = re.findall(phone_regex, text)
        return emails, phones, text

# Function to process a directory of CV files
def process_directory(directory):
    data = []
    for filename in os.listdir(directory):
        if filename.endswith('.pdf') or filename.endswith('.doc') or filename.endswith('.docx'):
            file_path = os.path.join(directory, filename)
            emails, phones, text = extract_info(file_path)
            data.append({'Filename': filename, 'Emails': emails, 'Phones': phones, 'Text': text})
    return data

# Specify the directory containing the CV files
cv_directory = (r"C:\Users\Nakul\Desktop\Sample2-20240406T093029Z-001\Sample2")

# Process the CV files and store the data
data = process_directory(cv_directory)

# Create a pandas DataFrame from the data
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
output_file = 'cv_data.xlsx'
df.to_excel(output_file, index=False)
print(f'Output saved to {output_file}')