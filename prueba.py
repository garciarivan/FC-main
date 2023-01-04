from docx import Document
from docx.enum.text import WD_COLOR_INDEX

# Crear un documento vacío
document = Document()

# Agregar una tabla al documento
table = document.add_table(rows=2, cols=2)

# Obtener la primera fila de la tabla
row = table.rows[0]

# Obtener las celdas de la primera fila
cell1 = row.cells[0]
cell2 = row.cells[1]

# Escribir dentro de las celdas
cell1.text = "Celda 1"
cell2.text = "Celda 2"

# Guardar el documento
document.save("table.docx")
