import os
import docx2txt
import re
import pandas as pd
import PyPDF2

# Define regex patterns to extract phone number and email
phone_pattern = re.compile(r'\d{10}|\d{3}-\d{3}-\d{4}|\d{3} \d{3} \d{4}|\d{5} \d{5}')
email_pattern = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
name_pattern = re.compile(r'([A-Z][a-z]+) ([A-Z][a-z]+)')
experience_pattern= re.compile(r'\b(\d+)\s+years?\b')



# Create an empty list to store the extracted information
info_list = []

# Loop through the .docx files in the folder
folder_path = (r"C:\Users\S.R SHALINI RAJENDAR\Downloads\OneDrive_2023-02-28\Embedded freshers_2-4yrs")
for filename in os.listdir(folder_path):
    if filename.endswith(".docx"):
        # Extract the text from the file
        file_path = os.path.join(folder_path, filename)
        text = docx2txt.process(file_path)
  
        name = text.split('\n')[0]   # Extract the name from the text (assuming the name is the first line)
  
        phone_match = phone_pattern.search(text) # Extract the phone number from the text using regex
        if phone_match:
            phone = phone_match.group()
        else:
            phone = ''

        email_match = email_pattern.search(text) # Extract the email address from the text using regex
        if email_match:
            email = email_match.group()
        else:
            email = ''
            
        experience_match = experience_pattern.search(text)
        if experience_match:
            experience=experience_match.group()
        else:
            experience = ''
            

        # Add the information to the list
        info_list.append([name, phone, email,experience])
      

    elif filename.endswith(".pdf"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'rb') as pdf_file:
                reader = PyPDF2.PdfReader(pdf_file)
                page = reader.pages[0]
                text = page.extract_text()
            # Extract the name from the text (assuming the name is the first line)
            name = text.split('\n')[0]
            
            
        

            # Extract the phone number from the text using regex
            phone_match = phone_pattern.search(text)
            if phone_match:
                phone = phone_match.group()
            else:
                phone = ''

            # Extract the email address from the text using regex
            email_match = email_pattern.search(text)
            if email_match:
                email = email_match.group()
            else:
                email = ''
            experience_match=experience_pattern.search(text)
            if experience_match:
                experience=experience_match.group()
            else:
                experience = ''


            # Add the information to the list
            info_list.append([name, phone, email,experience])

    # Convert the list to a Pandas DataFrame
df = pd.DataFrame(info_list, columns=['Name', 'Phone', 'Email','Experience'])

# Save the DataFrame to an Excel file
excel_path = (r"C:\Users\S.R SHALINI RAJENDAR\Downloads\Untitled spreadsheet (1).xlsx")
df.to_excel(excel_path, index=False)
