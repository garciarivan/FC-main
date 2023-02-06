'''from docx import Document
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
'''
from openpyxl import load_workbook, worksheet
filename = 'assets\listado_pacientes2.xlsx'
spreadsheet = load_workbook(filename)
sheet = spreadsheet.active

a = int(input('escribe numero de columna a mover: '))
worksheet.worksheet.Worksheet.move_range(cell_range="A1:C1", rows=0, cols=a)


print('Hecho')