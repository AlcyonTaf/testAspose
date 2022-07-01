# Import Aspose.Words for Python via .NET module
import aspose.words as aw



# Create and save a simple document
doc = aw.Document(".\TestWord\SansMacro.docx")

table = doc.get_child(aw.NodeType.TABLE, 4, True).as_table()
table5 = doc.get_child(aw.NodeType.TABLE, 5, True).clone(True).as_table()
allTables = doc.get_child_nodes(aw.NodeType.TABLE, True)
doc.get_child(aw.NodeType.TABLE, 5, True).remove()

row = table.rows[4]
cell = row.first_cell

row2 = aw.tables.Row(doc)

cell2 = aw.tables.Cell(doc)
cell2.append_child(aw.Paragraph(doc))
cell2.first_paragraph.append_child(aw.Run(doc, "test"))
cell3 = aw.tables.Cell(doc)
cell3.append_child(aw.Paragraph(doc))
cell3.first_paragraph.append_child(aw.Run(doc, "test"))
cell4 = aw.tables.Cell(doc)
cell4.append_child(aw.Paragraph(doc))
cell4.first_paragraph.append_child(aw.Run(doc, "test"))


row2.append_child(cell2)
row2.append_child(cell3)
row2.append_child(cell4)

table.rows.insert(4, row2)



#cell.cell_format.width = newwidth
# cell2.cell_format.width = newwidth
# cell3.cell_format.width = newwidth
#
# row.insert_after(cell2, cell)
# row.insert_after(cell3, cell2)






#
# for row in table.rows:
#     for cell in row.as_row().cells:
#         print(cell.to_string(aw.SaveFormat.TEXT))


doc.save('test2.docx')