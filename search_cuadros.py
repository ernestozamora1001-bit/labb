import openpyxl

file_path = r'c:\Users\emedrano\Downloads\Lab\Medrano Zamora.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb.active

print("=== SEARCHING FOR 'CUADRO' IN ALL CELLS (Rows 1-250) ===\n")

found = []
for row in range(1, 250):
    for col in range(1, 20):
        cell = ws.cell(row=row, column=col)
        val = cell.value
        if val:
            str_val = str(val).strip()
            if 'CUADRO' in str_val.upper():
                found.append((cell.coordinate, row, col, str_val))

if found:
    print(f"Found {len(found)} cells with 'CUADRO':")
    for coord, row, col, val in found:
        print(f"  {coord} (Row {row}): {val}")
else:
    print("No cells with 'CUADRO' found!")
    
    # Let's check row 46 specifically (from previous output)
    print("\n\nChecking Row 46 manually:")
    for col in range(1, 10):
        cell = ws.cell(row=46, column=col)
        print(f"  Col {col}: {repr(cell.value)}")

wb.close()
