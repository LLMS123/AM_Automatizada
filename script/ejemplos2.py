'''
import win32com
import docx
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

word = win32com.client.DispatchEx("Word.Application")
doc = word.Documents.Open()
doc.TablesOfContents(1).Update()
doc.Close(SaveChanges=True)
word.Quit()
'''

import win32com.client as win32

# Crear una instancia de la aplicación Word
word = win32.Dispatch('Word.Application')

# Abrir el archivo de Word y activar su ventana
file_path = r'C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\Am prueba v4.docx'
doc = word.Documents.Open(file_path)
word.Visible = False

# Actualizar la tabla de contenido
doc.TablesOfContents(1).Update()

#Section = win32.constants.wdSection
seccion = doc.Sections[3]
parrafo = seccion.Range.Paragraphs.First

# Verifica si el primer carácter del párrafo es un espacio en blanco y lo elimina si es el caso
if parrafo.Range.Characters.First.Text == '':
    parrafo.Range.Characters.First.Delete()
    
#first_section = doc.Sections(4) # obtener la primera sección del documento
#first_section.Range.Delete() # eliminar el contenido de la sección
#doc.Range(first_section.Range.End, first_section.Range.End).Delete()

#section_range = doc.Range(doc.Sections(4).Range.Start, doc.Sections(4).Range.End)
#section_range.Delete() # eliminar la sección del documento

doc.Close(SaveChanges=True) # guardar y cerrar el documento
#word.Quit() # cerrar la aplicación de Word

# Guardar y cerrar el archivo
#file_path=r'C:\Users\ANALISTAUP29\OneDrive - Ministerio de Educación\MINEDU_2022\GESTION DE LA INFORMACIÓN\UPP\Am Automatizada v2\AM_Automatizada\Am prueba v4.docx'
#doc.Save(file_path)
#doc.Close()

# Cerrar la aplicación Word
word.Quit()
