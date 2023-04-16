import pandas as pd
from docx import Document

# Load the Excel file into a pandas dataframe
df = pd.read_excel('D:/Coding/CreateWordFromExcel/addresses.xlsx', header=None)

# Loop through the rows in the dataframe
for index, row in df.iterrows():
    # Extract the ID, Address, and Coordinates from the row
    id = str(row[0]).replace('.pdf', '')
    address = row[1]
    coordinates = row[2]
    

    # Create a new Word document
    document = Document()

    # Add content to the Word document
    document.add_heading('Document Title', 0)
    document.add_paragraph('ID: ' + id)
    document.add_paragraph('Address: ' + address)
    document.add_paragraph('Coordinates: ' + coordinates)

    # Save the Word document with the ID and application form as the file name
    file_name = id + ' ' + 'application form' + '.docx'
    document.save(file_name)