import openpyxl

file_path = r'c:\Users\emedrano\Downloads\Lab\Medrano Zamora.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb.active

def check_cuadro(start_row, end_row, name):
    print(f"\n=== {name} (Rows {start_row}-{end_row}) ===")
    formulas_found = []
    empty_cells = []
    data_cells = []
    
    for row in range(start_row, end_row + 1):
        for col in range(1, 11):  # Check columns A-J
            cell = ws.cell(row=row, column=col)
            coord = cell.coordinate
            val = cell.value
            
            if val and isinstance(val, str) and val.startswith('='):
                formulas_found.append((coord, val))
            elif val is None or val == '':
                # Check if this cell might need a formula (based on context)
                pass
            else:
                data_cells.append((coord, val))
    
    if formulas_found:
        print(f"Formulas found: {len(formulas_found)}")
        for coord, formula in formulas_found[:10]:  # Show first 10
            print(f"  {coord}: {formula}")
        if len(formulas_found) > 10:
            print(f"  ... and {len(formulas_found) - 10} more")
    else:
        print("No formulas found!")
    
    return formulas_found

# Check each CUADRO
cuadros = [
    (46, 65, "CUADRO 1: GASTOS PREOPERATIVOS"),
    (67, 81, "CUADRO 2: INVERSION EN ACTIVO FIJO"),
    (82, 89, "CUADRO 3: PROGRAMA DE PRODUCCION"),
    (90, 105, "CUADRO 4: REQUERIMIENTO DE MATERIALES Y MANO DE OBRA"),
    (106, 134, "CUADRO 5: REQUERIMIENTOS DE CAPITAL DE TRABAJO"),
    (135, 156, "CUADRO 6"),
    (157, 200, "CUADRO 7: COSTOS DE OPERACION"),
]

for start, end, name in cuadros:
    check_cuadro(start, end, name)

wb.close()
