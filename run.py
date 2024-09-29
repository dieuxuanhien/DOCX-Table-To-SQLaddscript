from docx import Document

docx_path = "tables.docx"

doc = Document(docx_path)

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
    



scriptFile.close()
