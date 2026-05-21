import openpyxl

file_path = r'c:\Users\emedrano\Downloads\Lab\Medrano Zamora.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb['Fundamento del Pronóstico']

print("=== ANALYZING 'Fundamento del Pronóstico' SHEET ===\n")

# CUADRO locations based on previous search
cuadros = [
    (47, "CUADRO 1: GASTOS PREOPERATIVOS"),
    (68, "CUADRO 2: INVERSION EN ACTIVO FIJO"),
    (83, "CUADRO 3: PROGRAMA DE PRODUCCION"),
    (91, "CUADRO 4: REQUERIMIENTO DE MATERIALES Y MANO DE OBRA"),
    (107, "CUADRO 5: REQUERIMIENTOS DE CAPITAL DE TRABAJO"),
    (136, "CUADRO 6"),
    (158, "CUADRO 7: COSTOS DE OPERACION"),
]

# Estimate end of each cuadro
for i, (start_row, name) in enumerate(cuadros):
    end_row = cuadros[i+1][0] if i+1 < len(cuadros) else 250
    
    print(f"\n{'='*60}")
    print(f"{name} (Rows {start_row}-{end_row-1})")
    print(f"{'='*60}")
    
    # Count formulas in this range
    formula_count = 0
    empty_count = 0
    total_cells = 0
    
    for row in range(start_row, min(start_row + 25, end_row)):
        for col in range(1, 15):  # Columns A-N
            cell = ws.cell(row=row, column=col)
            total_cells += 1
            
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                formula_count += 1
            elif cell.value is None or cell.value == '':
                empty_count += 1
    
    print(f"Total cells checked: {total_cells}, Formulas: {formula_count}, Empty: {empty_count}")
    
    # Show first few rows of data
    print("\nFirst 10 rows:")
    for row in range(start_row, min(start_row + 12, end_row)):
        row_data = []
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            v = cell.value
            if v is not None and v != '':
                if isinstance(v, str) and v.startswith('='):
                    row_data.append(f"[{col}]{v[:30]}...")
                else:
                    row_data.append(f"[{col}]{v}")
        if row_data:
            print(f"  Row {row}: {' | '.join(row_data[:5])}")

wb.close()
