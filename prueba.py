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
import datetime

'''def fechado_archivo():
    #fecha_actual = datetime.datetime.now()    
    fecha_actual = datetime.date(2023, 2, 9)
    MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    if fecha_actual.weekday() == 0:
        duracion = datetime.timedelta(days=6)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 1:
        duracion = datetime.timedelta(days=5)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 1, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 2:
        duracion = datetime.timedelta(days=4)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 2, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 3:
        duracion = datetime.timedelta(days=3)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 3, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 4:
        duracion = datetime.timedelta(days=2)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 4, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 5:
        duracion = datetime.timedelta(days=1)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 5, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
    elif fecha_actual.weekday() == 6:
        duracion = datetime.timedelta(days=0)
        fecha_futura = fecha_actual + duracion
        return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - 6, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)

print(fechado_archivo())'''
'''def fechado_archivo():
    fecha_actual = datetime.datetime.now()
    MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    dias = {0: 6, 1: 5, 2: 4, 3: 3, 4: 2, 5: 1, 6: 0}
    duracion = datetime.timedelta(days=dias[fecha_actual.weekday()])
    fecha_futura = fecha_actual + duracion
    return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day - dias[fecha_actual.weekday()], fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
print(fechado_archivo())'''
'''from dateutil.relativedelta import relativedelta
import calendar

def fechado_archivo():
    #fecha_actual = datetime.datetime.now()
    fecha_actual = datetime.date(2023, 2, 9) 
    MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    dias_mes = calendar.monthrange(fecha_actual.year, fecha_actual.month)[1]
    fecha_futura = fecha_actual + datetime.timedelta(days=6)
    if fecha_futura.day > dias_mes:
        fecha_futura = fecha_futura.replace(month=fecha_futura.month+1, day=1)
    return 'LISTADO RECETAS SEMANAL DEL {} AL {} DE {} {}'.format(fecha_actual.day, fecha_futura.day, MESES[fecha_futura.month - 1].upper(), fecha_futura.year)
print(fechado_archivo())'''

'''#import calendar
MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
def fechado_archivo():
    fecha_actual = datetime.datetime.now()  
    dias_a_lunes = fecha_actual.weekday()
    dias_a_domingo = 6 - dias_a_lunes  
    fecha_lunes = fecha_actual - datetime.timedelta(days=dias_a_lunes)
    fecha_domingo = fecha_actual + datetime.timedelta(days=dias_a_domingo)
    return f'LISTADO RECETAS SEMANAL DEL {fecha_lunes.day} AL {fecha_domingo.day} DE {MESES[fecha_domingo.month -1].upper()} {fecha_actual.year}'
print(fechado_archivo())'''

from mailmerge import MailMerge
template = r'C:\Users\medico.RSD\Documents\FC-main\playground\template.docx'
document = MailMerge(template)

document.merge(map='Carmen Ballesteros'.upper(), lunes=str(27), domingo=str(5), mes='Marzo'.upper(), year=str(2023))
document.write(r'C:\Users\medico.RSD\Documents\FC-main\playground\prueba.docx')
print('HECHO!')