from docx import Document
import pyperclip

docx_path = "tables.docx"

doc = Document(docx_path)
# with open("script.txt", 'w') as file:
#         # This step creates the file or clears it if it already exists
#         pass  # No action needed, just opening in 'w' mode clears the file
scriptFile = open("script.txt",'w') 


for table in doc.tables:

    

    #get table name
    tableName = table.rows[0].cells[0].text.strip()


    #get headers
    headers = [cell.text.strip() for cell in table.rows[1].cells]

    scriptFile.write(f"-----------------------{tableName} DATA------------------------\n\n")
    #write tuples add
    for tuple in table.rows[2:]:
        values = [ f"'{cell.text.strip() }'"  for cell in tuple.cells]
        sqladd =f"INSERT INTO {tableName} ({",".join(headers)}) VALUES ({",".join(values)});"
        scriptFile.write(sqladd + "\n")


with open("script.txt", "r") as s:
    contents = s.read()
    pyperclip.copy(contents)


scriptFile.close()
