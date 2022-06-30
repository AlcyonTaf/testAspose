# Import Aspose.Words for Python via .NET module
import aspose.words as aw

# Create and save a simple document
doc = aw.Document(".\TestWord\SansMacro.docx")


allTables = doc.get_child_nodes(aw.NodeType.TABLE, True)

table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()


for row in table.rows:
    for cell in row.as_row().cells:
        print(cell.to_string(aw.SaveFormat.TEXT))