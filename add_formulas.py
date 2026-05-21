import openpyxl
from openpyxl.utils import get_column_letter

file_path = r'c:\Users\emedrano\Downloads\Lab\Medrano Zamora.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb['Fundamento del Pronóstico']

print("=== ADDING FORMULAS TO PENDING CUADROS ===\n")

# ============================================================
# CUADRO 4: REQUERIMIENTO DE MATERIALES Y MANO DE OBRA
# Rows 100-105 need formulas to pull from CUADRO 7
# ============================================================
print("CUADRO 4: Adding formulas for rows referencing CUADRO 7...")

# Row 100: REPUESTOS Y MANTENIMIENTO (CUADRO 7)
# Should reference CUADRO 7 row 169 (TOTAL repuestos)
ws['C100'] = '=C169'
ws['D100'] = '=D169'
ws['E100'] = '=E169'

# Row 101: GASTOS GENERALES (CUADRO 7)
# Should reference CUADRO 7 row 178 (TOTAL gastos indirectos)
ws['C101'] = '=C178'
ws['D101'] = '=D178'
ws['E101'] = '=E178'

# Row 103: ADMINISTRACION GENERAL Y SALARIOS (CUADRO 7)
ws['C103'] = '=C183'
ws['D103'] = '=D183'
ws['E103'] = '=E183'

# Row 104: GASTO DE VENTAS Y SALARIOS (CUADRO 7)
ws['C104'] = '=C186'
ws['D104'] = '=D186'
ws['E104'] = '=E186'

# Row 105: GASTO DE DISTRIBUCION Y SALARIOS (CUADRO 7)
ws['C105'] = '=C190'
ws['D105'] = '=D190'
ws['E105'] = '=E190'

print("  Added formulas for CUADRO 4 rows 100-105")

# ============================================================
# CUADRO 5: REQUERIMIENTOS DE CAPITAL DE TRABAJO
# Calculate working capital requirements
# ============================================================
print("\nCUADRO 5: Adding working capital calculation formulas...")

# Row 111: CUENTAS POR COBRAR
# Formula: (Annual Sales / 360) * Days
# Using CUADRO 3 production data and sales prices
ws['D111'] = '=($C$38*E88/360)*C111'  # Using max capacity production * price / 360 * days

# Row 114: Materiales locales - locales
# Based on CUADRO 4 material costs
ws['D114'] = '=(C95/360)*C114'

# Row 115: Materiales locales - importados
ws['D115'] = '=(C96/360)*C115'

# Row 116: Repuestos para mantenimiento
ws['D116'] = '=(C100/360)*C116'

# Row 117: Productos en proceso
# Average of materials + labor + services during production
ws['D117'] = '=((C95+C96+C97+C98)/2/360)*C117'

# Row 118: Producto terminado
ws['D118'] = '=(C95+C96+C97+C98)/360*C118'

# Row 120: E) GASTOS FINANCIEROS (existing formula: =+C111+C117+C118)
# This is already set

# PASIVO CIRCULANTE calculations
# Row 124: A. MATERIALES (pagados por adelantado = 0)
ws['E124'] = '=D124*0'

# Row 125: B. REPUESTOS (pagados por adelantado = 0)
ws['E125'] = '=D125*0'

# Row 126: C. MANO DE OBRA - pago al final de mes
ws['D126'] = '=(C97/360)*C126'

# Row 127: D. SERVICIOS - pago al final de mes
ws['D127'] = '=(C98/360)*C127'

# Row 129: E. GASTOS ADMINISTRATIVOS, VENTAS Y DISTRIBUCION
ws['D129'] = '=((C103+C104+C105)/360)*C129'

# Row 131: G. GASTOS FINANCIEROS
ws['D131'] = '=(D94+D95+D93)/360*C131'  # Interest payments from debt service

# Working Capital Requirement calculation
# Row 134: CAPITAL DE TRABAJO REQUERIDO
ws['B134'] = 'CAPITAL DE TRABAJO REQUERIDO'
ws['C134'] = '=SUM(D111:D118)-SUM(D124:D131)'

print("  Added formulas for CUADRO 5 working capital calculations")

# ============================================================
# CUADRO 6: COSTOS FIJOS Y VARIABLES
# This is a classification table - may need totals
# ============================================================
print("\nCUADRO 6: Adding cost classification formulas...")

# Add totals row at the end
ws['B153'] = 'TOTAL COSTOS FIJOS'
ws['C153'] = '=SUM(C142:C152)'
ws['D153'] = '=SUM(D142:D152)'

ws['B154'] = 'COSTO TOTAL'
ws['C154'] = '=C153'
ws['D154'] = '=D153'

print("  Added formulas for CUADRO 6 totals")

# Save the workbook
print("\n=== SAVING CHANGES ===")
wb.save(file_path)
print(f"Saved to: {file_path}")
wb.close()
print("\n=== DONE ===")
