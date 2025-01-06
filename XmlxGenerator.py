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
green_background = workbook.add_format({'bg_color': '#00FF00', 'font_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter'})
red_text = workbook.add_format({'font_color': '#FF0000', 'align': 'center', 'valign': 'vcenter'})
green_text = workbook.add_format({'font_color': '#00AA00', 'align': 'center', 'valign': 'vcenter'})
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
                    keysAnon.insert(0, key)
                elif key != "Name":
                    keysAnon.append(key)
            keys.append("Valid")
            keys.append("Nota")
            keysAnon.append("Valid")
            keysAnon.append("Nota")

        else:
            raise ValueError("El archivo JSON está vacío.")
except FileNotFoundError:
    print(f"El archivo {file_path} no se encontró.")
    exit(1)
except json.JSONDecodeError:
    print(f"Error al decodificar el archivo JSON {file_path}.")
    exit(1)

# Añadir encabezado
porcentaje = ["10%", "10%", "10%", "20%", "50%"]
worksheet.write_row(0, 0, keys, bold)
worksheet1.write_row(0, 0, keysAnon, bold)
worksheet.write_row(1, 1, porcentaje, bold)
worksheet1.write_row(1, 1, porcentaje, bold)

# Añadir datos
row = 2
rowFormula = row + 1
for entry in data:
    simplyId = entry["id"][1:5]

    # En la hoja "Notes amb nom"
    worksheet.write(row, 0, entry["Name"])
    worksheet.write(row, 1, entry["PR01"])
    worksheet.write(row, 2, entry["PR02"])
    worksheet.write(row, 3, entry["PR03"])
    worksheet.write(row, 4, entry["PR04"])
    worksheet.write(row, 5, entry["EX01"])
    worksheet.write(row, 6, entry["%Faltes"])

    # En la hoja "Notes anonimes", referenciamos las celdas de la hoja "Notes amb nom"
    worksheet1.write(row, 0, simplyId)
    worksheet1.write_formula(row, 1, f"='Notes amb nom'!B{row + 1}")  # Referencia a PR01
    worksheet1.write_formula(row, 2, f"='Notes amb nom'!C{row + 1}")  # Referencia a PR02
    worksheet1.write_formula(row, 3, f"='Notes amb nom'!D{row + 1}")  # Referencia a PR03
    worksheet1.write_formula(row, 4, f"='Notes amb nom'!E{row + 1}")  # Referencia a PR04
    worksheet1.write_formula(row, 5, f"='Notes amb nom'!F{row + 1}")  # Referencia a EX01
    worksheet1.write_formula(row, 6, f"='Notes amb nom'!G{row + 1}")  # Referencia a %Faltes

    # Validación: si más de un 20% de faltas o menos de 4 en el examen, no válido
    worksheet.write_formula(row, 7, f'=SI(Y(G{rowFormula}<=20, F{rowFormula}>=4), "Valid", "No valid")')
    worksheet1.write_formula(row, 7, f'=SI(Y(G{rowFormula}<=20, F{rowFormula}>=4), "Valid", "No valid")')

    # Fórmula para el cálculo de la nota final
    worksheet.write_formula(row, 8, f'B{row + 1}*0.10 + C{row + 1}*0.10 + D{row + 1}*0.10 + E{row + 1}*0.20 + F{row + 1}*0.50')
    worksheet1.write_formula(row, 8, f'B{row + 1}*0.10 + C{row + 1}*0.10 + D{row + 1}*0.10 + E{row + 1}*0.20 + F{row + 1}*0.50')

    # Condicional para cambiar el color del texto según la nota de la actividad (rojo si es menor que 5)
    worksheet.conditional_format(row, 1, row, 5, {'type': 'cell', 'criteria': '<', 'value': 5, 'format': red_text})
    worksheet1.conditional_format(row, 1, row, 5, {'type': 'cell', 'criteria': '<', 'value': 5, 'format': red_text})

    # Condicional para cambiar el color de fondo de la nota final (rojo si es menor a 5, verde si es mayor o igual a 7)
    worksheet.conditional_format(row, 8, row, 8, {'type': 'cell', 'criteria': '<', 'value': 5, 'format': red_background})
    worksheet1.conditional_format(row, 8, row, 8, {'type': 'cell', 'criteria': '<', 'value': 5, 'format': red_background})

    worksheet.conditional_format(row, 8, row, 8, {'type': 'cell', 'criteria': '>=', 'value': 7, 'format': green_background})
    worksheet1.conditional_format(row, 8, row, 8, {'type': 'cell', 'criteria': '>=', 'value': 7, 'format': green_background})

    rowFormula += 1
    row += 1

# Guardar el archivo
workbook.close()
print(f"Generated: '{filename}'")
