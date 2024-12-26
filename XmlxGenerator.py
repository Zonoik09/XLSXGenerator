#!/usr/bin/env python3
import xlsxwriter
import json

filename = "notes.xlsx"
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet("Notes amb nom")
worksheet1 = workbook.add_worksheet("Notes anonimes")

# Formats
bold = workbook.add_format({'bold': True})
italic = workbook.add_format({'italic': True})
centered = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
right_aligned = workbook.add_format({'align': 'right', 'valign': 'vcenter'})
red_background = workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter'})
green_text = workbook.add_format({'font_color': '#00AA00', 'align': 'right', 'valign': 'vcenter'})
bold_total = workbook.add_format({'bold': True, 'align': 'left'})

# Datos
file_path = "notes.json"

try:
    with open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
        if data:
            # Obtener las claves excluyendo 'id'
            keys = []
            for key in data[0].keys():
                if key != "id":
                    keys.append(key)

            # Obtener las claves excluyendo 'Name'
            keysAnon = []
            for key in data[0].keys():
                if key == "id":
                    keysAnon.insert(0,key)
                elif key != "Name":
                    keysAnon.append(key)

        else:
            raise ValueError("El archivo JSON está vacío.")
except FileNotFoundError:
    print(f"El archivo {file_path} no se encontró.")
    exit(1)
except json.JSONDecodeError:
    print(f"Error al decodificar el archivo JSON {file_path}.")
    exit(1)

# Añadir encabezado
porcentaje = ["10%","10%","10%","20%","50%"]

worksheet.write_row(0, 0, keys, bold)
worksheet1.write_row(0,0,keysAnon,bold)
worksheet.write_row(1,1,porcentaje,bold)
worksheet1.write_row(1,1,porcentaje,bold)




# Añadir datos
row = 2

for entry in data:
    simplyId = entry["id"][1:5]
    worksheet.write(row, 0, entry["Name"])
    worksheet.write(row, 1, entry["PR01"])
    worksheet.write(row, 2, entry["PR02"])
    worksheet.write(row, 3, entry["PR03"])
    worksheet.write(row, 4, entry["PR04"])
    worksheet.write(row, 5, entry["EX01"])
    worksheet.write(row, 6, entry["%Faltes"])
    worksheet1.write(row, 0, simplyId)
    worksheet1.write(row, 1, entry["PR01"])
    worksheet1.write(row, 2, entry["PR02"])
    worksheet1.write(row, 3, entry["PR03"])
    worksheet1.write(row, 4, entry["PR04"])
    worksheet1.write(row, 5, entry["EX01"])
    worksheet1.write(row, 6, entry["%Faltes"])
    row += 1

# Guardar el archivo
workbook.close()
print(f"Generated: '{filename}'")


# Pag1 ! A1