import pandas as pd
from docx import Document
from docxtpl import DocxTemplate

# Load the Excel file into a pandas dataframe
df = pd.read_excel('D:/Coding/CreateWordFromExcel/addresses.xlsx', header=None)

template = Document('D:/Coding/CreateWordFromExcel/application_form.docx')
# Loop through the rows in the dataframe
for index, row in df.iterrows():
    document = Document('D:/Coding/CreateWordFromExcel/application_form.docx')
    # Extract the ID, Address, and Coordinates from the row
    id = str(row[0]).replace('.pdf', '')
    address = row[1]
    coordinates = row[2]
    
    # Create a new Word document
    table = document.tables[0]
    cell_1 = table.cell(1, 5)  # Access the second cell in the first row
    cell_2 = table.cell(3,2 )  # Access the second cell in the second row
    cell_3 = table.cell(2,2)  # Access the second cell in the third row
    
    cell_1.add_paragraph(id)
    cell_2.add_paragraph(address)
    cell_3.add_paragraph(coordinates)

    # Save the Word document with the ID and application form as the file name
    filename = f"{id} application form.docx"
    document.save(f"output_folder/{filename}")
