#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
#from openpyxl import Workbook
import openpyxl
from data import change_column_fill, copy_format_column, format_condition_iter2, insert_column, remove_rows, subtract_two_column, tick_red_cell, unmerge_cells
#change_column_AdditionalNotes, change_column_TeamRequestComment1, delete_rows_LeadPracticeArea, delete_rows_TeamRequestStatu, find_column_Team, find_column_lead,
from miargparse import parser
from openpyxl.styles import PatternFill 

args = parser.parse_args()
print(args)

wb = openpyxl.load_workbook(args.ruta_input)
ws = wb['Retain Report']

#Quita el formato de celdas compartidas
unmerge_cells(ws)

#Elimina las filas de la (1 a la 11)
remove_rows(ws,1,11)

#Cambia a todas la celdas les pone color blanco (quita el amarillo)
change_column_fill(ws)

#Metodo que da formato a la columna que se ha creado según criterio  
color_yellow = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type = "solid")
format_condition_iter2(ws, colNr= 9, condition="CRITICA", color=color_yellow)

# Metodo que inserta una columna nueva detrás de la columna H poniendo la cabecera y copiando el formato de la columna H
insert_column(ws, colNr=9, headerRow=1, headerVal='FTES. Pdtes.')

#Metodo que da formato a la columna que se ha creado 
copy_format_column(ws, colNr= 9)

#Sumamos las columnas G y H --------------------------->
subtract_two_column(ws)

#Poner la columna en rojo si no se ha podido restar (G Y H)
color_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type = "solid")
tick_red_cell(ws, color=color_red, colNr= 9)

#Borrar Filas columna R -- Eliminar las que no sean CCA_SCE_ES
#delete_rows_LeadPracticeArea(ws, find_column_lead(ws))

#Borrar filas columna C -- Eliminar las filas que sean igual a Draf
#delete_rows_TeamRequestStatu(ws,find_column_Team(ws))

#Busca el nombre de la cabecera F y lo cambia por CRITICIDAD 
#change_column_AdditionalNotes(ws)

#Busca si la columna nueva 'FTES. Pdtes = 0, poner en la columna CRITICIDAD el valor ‘CUBIERTA’
#change_column_criticality(ws, find_column_criticality)

#Busca el nombre de la cabecera O y lo cambia por CLIENTE
#change_column_TeamRequestComment1(ws)




# Save the workbook to the output file
wb.save(args.ruta_output)


