import openpyxl

file_path = r'c:\Users\emedrano\Downloads\Lab\Medrano Zamora.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb.active

print("=== CUADRO 3: PROGRAMA DE PRODUCCION (Rows 82-95) ===")
print("\nRow by Row Analysis:\n")

for row in range(82, 96):
    print(f"Row {row}:")
    for col in range(1, 11):
        cell = ws.cell(row=row, column=col)
        coord = cell.coordinate
        val = cell.value
        if val is not None and val != '':
            print(f"  {coord}: {val}")
    print()

print("\n=== Looking for patterns to understand what formulas are needed ===")

# Check CUADRO 4 to understand the connection
print("\n=== CUADRO 4 (Rows 90-105) - Checking for references to CUADRO 3 ===")
for row in range(90, 106):
    for col in range(1, 11):
        cell = ws.cell(row=row, column=col)
        val = cell.value
        if val and isinstance(val, str) and 'CUADRO' in val.upper():
            print(f"  {cell.coordinate}: {val}")

wb.close()
