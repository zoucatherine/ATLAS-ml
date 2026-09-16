from docx import Document

# Open the DOCX file
doc = Document("example.docx")

print("TEXT:")
print("-" * 30)

# Extract paragraphs
for paragraph in doc.paragraphs:
    if paragraph.text.strip():
        print(paragraph.text)

print("\nTABLES:")
print("-" * 30)

# Extract tables
for table_num, table in enumerate(doc.tables, start=1):
    print(f"\nTable {table_num}:")

    for row in table.rows:
        row_data = [cell.text for cell in row.cells]
        print(row_data)