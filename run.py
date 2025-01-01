from docx import Document
import pyperclip

docx_path = "tables.docx"

doc = Document(docx_path)
with open("script.txt", 'w') as file:
        # This step creates the file or clears it if it already exists
        pass  # No action needed, just opening in 'w' mode clears the file
scriptFile = open("script.txt",'w', encoding='utf-8') 

for table in doc.tables:
    #get table name
    tableName = table.rows[0].cells[0].text.split(" ")[0]

    #get headers
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    headers[0] = headers[0].split(" ")[1]
    scriptFile.write(f"-----------------------{tableName} DATA------------------------\n\n")
    #write tuples add
    for tuple in table.rows[1:]:
        values = []
        for cell in tuple.cells:
            cell_text = cell.text.strip()
            if cell_text == '0'  or ( cell_text.isnumeric() and not cell_text.startswith('0')):
                values.append(cell_text)
            else:
                values.append(f"'{cell_text}'")
        sqladd = f"INSERT INTO {tableName} ({','.join(headers)}) VALUES ({','.join(values)});"
        scriptFile.write(sqladd + "\n")

scriptFile.close()