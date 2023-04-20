import pandas as pd
from docx import Document
from docxtpl import DocxTemplate
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Load the Excel file into a pandas dataframe
df = pd.read_excel(r'D:\CODING\PYTHON\BTK_BAZ_İSTASYON\Guv_Sert_Parser\SiteID_AdresKoord.xlsx', header=None)

template = Document(r'D:\CODING\PYTHON\BTK_BAZ_İSTASYON\CreateWordFromExcel\application_form.docx')
# Loop through the rows in the dataframe
for index, row in df.iterrows():
    document = Document(r'D:\CODING\PYTHON\BTK_BAZ_İSTASYON\CreateWordFromExcel\application_form.docx')
    # Extract the ID, Address, and Coordinates from the row
    id = str(row[0]).replace('.pdf', '')
    id = str(row[0]).replace('SER', '')

    address = row[1]
    coordinates = row[2]
    cord_len = (len(coordinates))
    print(type(cord_len))
    longitutde =coordinates[0:cord_len//2].strip()
    latitude = coordinates[((cord_len//2)+1):cord_len].strip()
    
    # Create a new Word document
    table = document.tables[0]
    cell_1 = table.cell(1, 5)  # add the site ID to template
    cell_2 = table.cell(3,2 )  # add site address to the template
    cell_3 = table.cell(2,2)  # add site longitutede coordinate to the template
    cell_4 = table.cell(2,3)
    cell_5 = table.cell(6,4)
    
    cell_1.add_paragraph(id)
    cell_2.add_paragraph(address)
    cell_3.add_paragraph(longitutde)
    cell_4.add_paragraph(latitude)
    cell_5.add_paragraph(address)
    
    
    font_style = 'Calibri'
    font_size = 9
    
    # Define alignment
    alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    
    # Add the ID to the cell with the specified font and alignment
    p1 = cell_1.paragraphs[0]
    p1.add_run(id).font.name = font_style
    p1.add_run(' ').font.name = font_style
    p1.add_run('').font.name = font_style
    p1.alignment = alignment
    p1.runs[0].font.size = font_size

    # Add the address to the cell with the specified font and alignment
    p2 = cell_2.paragraphs[0]
    p2.add_run(address).font.name = font_style
    p2.alignment = alignment
    p2.runs[0].font.size = font_size

    # Add the coordinates to the cell with the specified font and alignment
    p3 = cell_3.paragraphs[0]
    p3.add_run(coordinates).font.name = font_style
    p3.alignment = alignment
    p3.runs[0].font.size = font_size
        
    p4 = cell_4.paragraphs[0]
    
    # Save the Word document with the ID and application form as the file name
    filename = f"{id} Başvuru Formu.docx"
    document.save(f"output_folder/{filename}")
