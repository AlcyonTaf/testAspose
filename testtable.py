# Import Aspose.Words for Python via .NET module
import aspose.words as aw
import aspose.words.tables as tables # <module 'aspose.words.tables'>


# Create and save a simple document
doc = aw.Document(".\TestWord\simpletableau.docx")

allTables = doc.get_child_nodes(aw.NodeType.TABLE, True)

table = doc.get_child(aw.NodeType.TABLE, 0, True).as_table()

builder = aw.DocumentBuilder(doc)

# Start building the table.
table2 = builder.start_table()
builder.insert_cell()
builder.write("Row 1, Cell 1 Content.")

# Build the second cell.
builder.insert_cell()
builder.write("Row 1, Cell 2 Content.")

# Call the following method to end the row and start a new row.
builder.end_row()

# Build the first cell of the second row.
builder.insert_cell()
builder.write("Row 2, Cell 1 Content")

# Build the second cell.
builder.insert_cell()
builder.write("Row 2, Cell 2 Content.")
builder.end_row()

# Signal that we have finished building the table.
builder.end_table()




while table.has_child_nodes:
    table2.rows.add(table.first_row)

# test = table.preferred_width.value
#
# print(test)
#
# clone = table.last_row.clone(True).as_row()
# for cell in clone.cells:
#     cell = cell.as_cell()
#     cell.remove_all_children()
#
# table.append_child(clone)

# row = aw.tables.Row(doc)
# row.row_format.allow_break_across_pages = True
# table.append_child(row)
# table.auto_fit(aw.tables.AutoFitBehavior.FIXED_COLUMN_WIDTHS)
# cell = aw.tables.Cell(doc)
# #cell.cell_format.shading.background_pattern_color = drawing.Color.light_blue
# cell.cell_format.width = 300
# cell.append_child(aw.Paragraph(doc))
# cell.first_paragraph.append_child(aw.Run(doc, "Row 1, Cell 1 Text"))
#
# row.append_child(cell)

for row in table.rows:
    for cell in row.as_row().cells:
        print(cell.to_string(aw.SaveFormat.TEXT))


doc.save(('testtable.docx'))
