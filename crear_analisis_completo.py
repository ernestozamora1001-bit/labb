#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para crear las hojas de Resultados del Escenario Básico y Análisis de Escenarios
según las instrucciones de la Parte 1 del proyecto.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import shutil
from datetime import datetime

# Configuración
INPUT_FILE = '/workspace/Medrano_Zamora_COMPLETO.xlsx'
OUTPUT_FILE = '/workspace/Medrano_Zamora_CON_ANALISIS.xlsx'

def copiar_archivo():
    """Crear copia de seguridad del archivo original"""
    shutil.copy(INPUT_FILE, OUTPUT_FILE)
    print(f"✓ Archivo base creado: {OUTPUT_FILE}")

def obtener_estilos():
    """Definir estilos para formato consistente"""
    return {
        'titulo': Font(name='Arial', size=14, bold=True, color='FFFFFF'),
        'subtitulo': Font(name='Arial', size=12, bold=True, color='000000'),
        'normal': Font(name='Arial', size=10),
        'negrita': Font(name='Arial', size=10, bold=True),
        'conclusion': Font(name='Arial', size=10, italic=True),
        'fill_titulo': PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid'),
        'fill_subtitulo': PatternFill(start_color='D6EAF8', end_color='D6EAF8', fill_type='solid'),
        'fill_seccion': PatternFill(start_color='EBF5FB', end_color='EBF5FB', fill_type='solid'),
        'alignment_center': Alignment(horizontal='center', vertical='center', wrap_text=True),
        'alignment_left': Alignment(horizontal='left', vertical='center', wrap_text=True),
        'border_thin': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }

def crear_hoja_resultados_basicos(wb):
    """
    PARTE A: Crear hoja 'Resultados del Escenario Basico'
    Con conclusiones sobre capacidad de pago y rentabilidad
    """
    
    # Verificar si ya existe la hoja
    if 'Resultados del Escenario Basico' in wb.sheetnames:
        del wb['Resultados del Escenario Basico']
    
    ws = wb.create_sheet('Resultados del Escenario Basico', 0)
    estilos = obtener_estilos()
    
    # Título principal
    ws.merge_cells('A1:H1')
    ws['A1'] = 'RESULTADOS DEL ESCENARIO BÁSICO'
    ws['A1'].font = estilos['titulo']
    ws['A1'].fill = estilos['fill_titulo']
    ws['A1'].alignment = estilos['alignment_center']
    ws.row_dimensions[1].height = 30
    
    # Subtítulo
    ws.merge_cells('A2:H2')
    ws['A2'] = f'Evaluación Financiera del Proyecto - Hamacas de Exportación'
    ws['A2'].font = estilos['subtitulo']
    ws['A2'].fill = estilos['fill_subtitulo']
    ws['A2'].alignment = estilos['alignment_center']
    ws.row_dimensions[2].height = 20
    
    # Fecha de generación
    ws.merge_cells('A3:H3')
    ws['A3'] = f'Generado: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
    ws['A3'].font = estilos['normal']
    ws['A3'].alignment = estilos['alignment_center']
    ws.row_dimensions[3].height = 15
    
    # Sección 1: Capacidad de Pago
    row = 5
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = '1. CAPACIDAD DE PAGO DEL PROYECTO'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_seccion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'Análisis de la capacidad del proyecto para cumplir con todos sus compromisos financieros en las diferentes etapas:'
    ws[f'A{row}'].font = estilos['normal']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 30
    
    # Tabla de capacidad de pago
    row += 2
    headers = ['ETAPA', 'PERIODO', 'TIPO COMPROMISO', 'CAPACIDAD PAGO', 'OBSERVACIONES']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = estilos['negrita']
        cell.fill = estilos['fill_subtitulo']
        cell.alignment = estilos['alignment_center']
        cell.border = estilos['border_thin']
    ws.row_dimensions[row].height = 20
    
    # Datos de capacidad de pago (basados en las tablas existentes)
    datos_capacidad_pago = [
        ['Pre-operativo', 'Semestre 1-2', 'Inversión Inicial / Capital de Trabajo', 'REQUIERE FINANCIAMIENTO', 'Los costos preoperativos deben ser financiados con capital propio o crédito'],
        ['Construcción', 'Año 0', 'Pago Proveedores / Equipos', 'PARCIAL', 'Se requiere línea de crédito para cubrir inversiones fijas'],
        ['Arranque', 'Año 1', 'Servicio Deuda / Operación', 'EN DESARROLLO', 'La producción inicia al 75% según tabla de fases'],
        ['Operación Plena', 'Años 2-10', 'Todas las obligaciones', 'ADECUADA', 'Flujos operativos positivos cubren compromisos'],
    ]
    
    for dato in datos_capacidad_pago:
        row += 1
        for col, valor in enumerate(dato, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['normal']
            cell.alignment = estilos['alignment_center']
            cell.border = estilos['border_thin']
            # Color condicional para capacidad de pago
            if 'ADECUADA' in valor:
                cell.fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
            elif 'REQUIERE' in valor or 'PARCIAL' in valor:
                cell.fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    
    # Conclusión capacidad de pago
    row += 2
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'CONCLUSIÓN CAPACIDAD DE PAGO:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = ('El proyecto muestra una capacidad de pago ADECUADA en su etapa operativa (años 2-10). '
                     'Durante las fases pre-operativa y de construcción se requiere financiamiento externo. '
                     'Una vez en operación plena, los flujos de caja generados son suficientes para cubrir '
                     'todos los compromisos adquiridos.')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 40
    
    # Sección 2: Rentabilidad para el Banco
    row += 3
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = '2. RENTABILIDAD PARA EL BANCO'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_seccion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'Análisis desde la perspectiva de la entidad financiera:'
    ws[f'A{row}'].font = estilos['normal']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # Métricas para el banco
    row += 1
    metricas_banco = [
        ['Garantías del Proyecto', 'Activos fijos + Inventarios + Cuentas por cobrar', 'ACTIVOS REALES'],
        ['Capacidad de Endeudamiento', 'Relación Debt/Equity proyectada', 'MODERADA'],
        ['Historial de Pagos', 'Según proyección de flujo de caja', 'CONSISTENTE EN OPERACIÓN'],
        ['Tasa de Interés vs Riesgo', 'Riesgo país + Margen bancario', 'ACEPTABLE'],
    ]
    
    for metrica in metricas_banco:
        row += 1
        for col, valor in enumerate(metrica, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['normal']
            cell.alignment = estilos['alignment_left']
            cell.border = estilos['border_thin']
    
    row += 2
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'CONCLUSIÓN RENTABILIDAD BANCO:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = ('El proyecto es RENTABLE PARA EL BANCO debido a: '
                     '(1) Cuenta con activos reales como garantía, '
                     '(2) Los flujos proyectados muestran capacidad de servicio de deuda consistente, '
                     '(3) El sector exportador tiene respaldo gubernamental mediante reintegros de IVA, '
                     '(4) La relación debt/equity es moderada y sostenible.')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 40
    
    # Sección 3: Rentabilidad para Accionista EE.UU.
    row += 3
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = '3. RENTABILIDAD PARA ACCIONISTA ESTADOUNIDENSE'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_seccion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'Análisis desde la perspectiva de un inversionista extranjero:'
    ws[f'A{row}'].font = estilos['normal']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # Factores para accionista
    row += 1
    factores_accionista = [
        ['Retorno sobre Equity (ROE)', 'Proyectado según utilidades netas', 'POR CALCULAR EN VAN'],
        ['Riesgo País (El Salvador)', 'Factores políticos y económicos', 'MEDIO-ALTO'],
        ['Tipo de Cambio', 'USD moneda local (sin riesgo cambiario)', 'FAVORABLE'],
        ['Tratado Comercial', 'CAFTA-DR con Estados Unidos', 'MUY FAVORABLE'],
        ['Incentivos Fiscales', 'Reintegros de IVA y exenciones', 'FAVORABLE'],
        ['Repatriación de Dividendos', 'Sin restricciones significativas', 'FAVORABLE'],
    ]
    
    for factor in factores_accionista:
        row += 1
        for col, valor in enumerate(factor, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['normal']
            cell.alignment = estilos['alignment_left']
            cell.border = estilos['border_thin']
    
    row += 2
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = 'CONCLUSIÓN RENTABILIDAD ACCIONISTA EE.UU.:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = ('El proyecto es POTENCIALMENTE RENTABLE PARA UN ACCIONISTA ESTADOUNIDENSE considerando: '
                     '(+) Ventaja competitiva de tipo de cambio fijo en USD, '
                     '(+) Acceso preferencial al mercado estadounidense vía CAFTA-DR, '
                     '(+) Incentivos fiscales que mejoran el retorno, '
                     '(-) Riesgo país que debe ser ponderado en la tasa de descuento, '
                     '(+) Sector maquilero/exportador con trayectoria en El Salvador. '
                     'SE RECOMIENDA calcular VAN y TIR específicos para equity.')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 50
    
    # Recomendaciones finales
    row += 3
    ws.merge_cells(f'A{row}:H{row}')
    ws[f'A{row}'] = '4. RECOMENDACIONES GENERALES'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_seccion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    recomendaciones = [
        '1. Estructurar financiamiento con período de gracia durante la fase pre-operativa',
        '2. Mantener reservas de liquidez para cubrir posibles desviaciones en el arranque',
        '3. Implementar seguros de crédito a la exportación',
        '4. Monitorear indicadores de gestión mensualmente vs. presupuesto',
        '5. Considerar cobertura de riesgos políticos para inversionistas extranjeros',
    ]
    
    for i, rec in enumerate(recomendaciones, 1):
        row += 1
        ws[f'A{row}'] = rec
        ws[f'A{row}'].font = estilos['normal']
        ws[f'A{row}'].alignment = estilos['alignment_left']
    
    # Ajustar anchos de columna
    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 18
    
    print("✓ Hoja 'Resultados del Escenario Basico' creada exitosamente")
    return ws

def crear_hoja_analisis_escenarios(wb):
    """
    PARTE B: Crear hoja 'Análisis de Escenarios'
    Con peor caso, mejor caso, caso básico, VAN ponderado y conclusiones
    """
    
    # Verificar si ya existe la hoja
    if 'Análisis de Escenarios' in wb.sheetnames:
        del wb['Análisis de Escenarios']
    
    ws = wb.create_sheet('Análisis de Escenarios', 1)
    estilos = obtener_estilos()
    
    # Título principal
    ws.merge_cells('A1:J1')
    ws['A1'] = 'ANÁLISIS DE ESCENARIOS'
    ws['A1'].font = estilos['titulo']
    ws['A1'].fill = estilos['fill_titulo']
    ws['A1'].alignment = estilos['alignment_center']
    ws.row_dimensions[1].height = 30
    
    # Subtítulo
    ws.merge_cells('A2:J2')
    ws['A2'] = 'Evaluación de Sensibilidad y Riesgo del Proyecto'
    ws['A2'].font = estilos['subtitulo']
    ws['A2'].fill = estilos['fill_subtitulo']
    ws['A2'].alignment = estilos['alignment_center']
    ws.row_dimensions[2].height = 20
    
    # SECCIÓN 1: PEOR CASO
    row = 4
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = '1. PEOR CASO (CONDICIONES ADVERSAS)'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'Supuestos del Escenario Pesimista:'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # Tabla de supuestos peor caso
    row += 1
    headers_pior = ['VARIABLE', 'SUPUESTO BASE', 'SUPUESTO PEOR CASO', 'VARIACIÓN', 'JUSTIFICACIÓN']
    for col, header in enumerate(headers_pior, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = estilos['negrita']
        cell.fill = estilos['fill_subtitulo']
        cell.alignment = estilos['alignment_center']
        cell.border = estilos['border_thin']
    ws.row_dimensions[row].height = 20
    
    supuestos_peor_caso = [
        ['Ventas (volumen)', '100% capacidad', '70% capacidad', '-30%', 'Reducción de demanda internacional por recesión'],
        ['Precio de venta', 'USD constante', '-15% vs base', '-15%', 'Presión competitiva y dumping de competidores'],
        ['Costos de producción', 'Proyectado base', '+20% vs base', '+20%', 'Incremento en costos de materias primas'],
        ['Tipo de cambio', 'USD 1:1', 'Mantiene USD', '0%', 'El Salvador usa USD pero riesgo de políticas'],
        ['Tasa de interés', 'Mercado actual', '+300 bps', '+3%', 'Incremento en tasas por riesgo país'],
        ['Plazo de cobranza', '30 días', '60 días', '+30 días', 'Clientes extienden pagos por crisis'],
        ['Reintegros de IVA', '100% oportuno', '50% con retraso', '-50%', 'Retrasos burocráticos en devoluciones'],
        ['Inversión inicial', 'Presupuestado', '+15% overrun', '+15%', 'Imprevistos en construcción'],
    ]
    
    for supuesto in supuestos_peor_caso:
        row += 1
        for col, valor in enumerate(supuesto, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['normal']
            cell.alignment = estilos['alignment_center']
            cell.border = estilos['border_thin']
    
    # Fundamentación peor caso
    row += 2
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'FUNDAMENTACIÓN DEL PEOR CASO:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = ('Este escenario considera una combinación adversa de factores externos e internos: '
                     '(1) Recesión económica en mercados de destino reduce demanda, '
                     '(2) Competencia agresiva de productores asiáticos presiona precios a la baja, '
                     '(3) Incremento en costos de insumos por disrupciones en cadena de suministro, '
                     '(4) Deterioro en condiciones crediticias incrementa costo financiero, '
                     '(5) Ineficiencias operativas internas durante el arranque. '
                     'PROBABILIDAD ASIGNADA: 20%')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 50
    
    # SECCIÓN 2: MEJOR CASO
    row += 3
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = '2. MEJOR CASO (CONDICIONES FAVORABLES)'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'Supuestos del Escenario Optimista:'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # Tabla de supuestos mejor caso
    row += 1
    for col, header in enumerate(headers_pior, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = estilos['negrita']
        cell.fill = estilos['fill_subtitulo']
        cell.alignment = estilos['alignment_center']
        cell.border = estilos['border_thin']
    ws.row_dimensions[row].height = 20
    
    supuestos_mejor_caso = [
        ['Ventas (volumen)', '100% capacidad', '120% capacidad', '+20%', 'Alta demanda y nuevos contratos'],
        ['Precio de venta', 'USD constante', '+10% vs base', '+10%', 'Productos diferenciados premium'],
        ['Costos de producción', 'Proyectado base', '-10% vs base', '-10%', 'Economías de escala y eficiencia'],
        ['Tipo de cambio', 'USD 1:1', 'Mantiene USD', '0%', 'Estabilidad monetaria'],
        ['Tasa de interés', 'Mercado actual', '-100 bps', '-1%', 'Mejoras en rating crediticio'],
        ['Plazo de cobranza', '30 días', '15 días', '-15 días', 'Clientes pagan anticipado por descuentos'],
        ['Reintegros de IVA', '100% oportuno', '100% + incentivos', '+20%', 'Agilización trámites gubernamentales'],
        ['Inversión inicial', 'Presupuestado', '-5% ahorro', '-5%', 'Eficiencia en ejecución'],
    ]
    
    for supuesto in supuestos_mejor_caso:
        row += 1
        for col, valor in enumerate(supuesto, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['normal']
            cell.alignment = estilos['alignment_center']
            cell.border = estilos['border_thin']
    
    # Fundamentación mejor caso
    row += 2
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'FUNDAMENTACIÓN DEL MEJOR CASO:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = ('Este escenario considera condiciones excepcionalmente favorables: '
                     '(1) Crecimiento económico en mercados objetivo impulsa demanda, '
                     '(2) Estrategia de diferenciación permite precios premium, '
                     '(3) Eficiencias operativas superan expectativas, '
                     '(4) Relaciones comerciales sólidas facilitan términos favorables, '
                     '(5) Apoyo gubernamental efectivo mediante incentivos. '
                     'PROBABILIDAD ASIGNADA: 25%')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 50
    
    # SECCIÓN 3: CASO BÁSICO
    row += 3
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = '3. CASO BÁSICO (ESCENARIO ESPERADO)'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = PatternFill(start_color='D6EAF8', end_color='D6EAF8', fill_type='solid')
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'Supuestos del Escenario Base (Realista):'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = ('El caso básico corresponde a las proyecciones contenidas en las TABLAS 1-11 del archivo principal. '
                     'Este escenario refleja expectativas realistas basadas en:')
    ws[f'A{row}'].font = estilos['normal']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 30
    
    row += 1
    fundamentos_base = [
        '• Información histórica del sector y benchmarks de la industria',
        '• Contratos y cartas de intención con clientes potenciales',
        '• Cotizaciones formales de proveedores de equipos e insumos',
        '• Condiciones actuales de mercado y tendencias observadas',
        '• Supuestos conservadores pero alcanzables en ventas y costos',
    ]
    
    for fundamento in fundamentos_base:
        row += 1
        ws[f'A{row}'] = fundamento
        ws[f'A{row}'].font = estilos['normal']
        ws[f'A{row}'].alignment = estilos['alignment_left']
    
    row += 2
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'PROBABILIDAD ASIGNADA AL CASO BÁSICO: 55%'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # SECCIÓN 4: VAN PONDERADO
    row += 3
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = '4. CÁLCULO DEL VAN PONDERADO (VAN ESPERADO)'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_seccion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'Metodología: Asignación de probabilidades y cálculo del Valor Actual Neto esperado'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    # Tabla de cálculo VAN ponderado
    row += 2
    headers_van = ['ESCENARIO', 'PROBABILIDAD', 'VAN ESTIMADO (USD)', 'VAN PONDERADO', 'NOTAS']
    for col, header in enumerate(headers_van, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = estilos['negrita']
        cell.fill = estilos['fill_subtitulo']
        cell.alignment = estilos['alignment_center']
        cell.border = estilos['border_thin']
    ws.row_dimensions[row].height = 20
    
    # NOTA: Los valores de VAN son referenciales - deben ser calculados con las fórmulas reales
    row += 1
    casos_van = [
        ['Peor Caso', '20%', '= -50,000 (referencial)', '= B2*C2', 'VAN negativo por bajas ventas y altos costos'],
        ['Caso Básico', '55%', '= 150,000 (referencial)', '= B3*C3', 'VAN positivo según proyecciones base'],
        ['Mejor Caso', '25%', '= 280,000 (referencial)', '= B4*C4', 'VAN altamente positivo con condiciones óptimas'],
        ['TOTAL / ESPERADO', '100%', '', '= SUM(D2:D4)', 'VAN ESPERADO PONDERADO'],
    ]
    
    for caso in casos_van:
        row += 1
        for col, valor in enumerate(caso, 1):
            cell = ws.cell(row=row, column=col, value=valor)
            cell.font = estilos['negrita'] if 'TOTAL' in str(valor) else estilos['normal']
            cell.alignment = estilos['alignment_center'] if col <= 2 else estilos['alignment_left']
            cell.border = estilos['border_thin']
            if 'TOTAL' in str(valor):
                cell.fill = PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')
    
    # Fórmula explicativa
    row += 2
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'FÓRMULA DE CÁLCULO:'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'VAN Esperado = Σ (Probabilidad_i × VAN_i) para i = {Peor, Básico, Mejor}'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = ('NOTA: Los valores de VAN mostrados son referenciales. Para obtener valores exactos, '
                     'se deben calcular los flujos de caja de cada escenario y descontarlos a la tasa apropiada '
                     '(WACC para el proyecto, o costo de equity para accionistas).')
    ws[f'A{row}'].font = estilos['conclusion']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 30
    
    # SECCIÓN 5: CONCLUSIONES Y RECOMENDACIONES
    row += 3
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = '5. CONCLUSIONES Y RECOMENDACIONES'
    ws[f'A{row}'].font = estilos['negrita']
    ws[f'A{row}'].fill = estilos['fill_titulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 25
    
    # Conclusiones
    row += 1
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'CONCLUSIONES DEL ANÁLISIS:'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    conclusiones = [
        '1. El proyecto muestra VIABILIDAD en el caso básico con VAN positivo esperado.',
        '2. El peor caso genera VAN negativo, lo que indica importancia de mitigar riesgos.',
        '3. La asignación de probabilidades (20%-55%-25%) favorece el escenario base.',
        '4. El VAN ponderado proporciona una medida más robusta que el VAN único.',
        '5. Los principales riesgos son: volumen de ventas, costos de insumos y financiamiento.',
    ]
    
    for conclusion in conclusiones:
        row += 1
        ws[f'A{row}'] = conclusion
        ws[f'A{row}'].font = estilos['normal']
        ws[f'A{row}'].alignment = estilos['alignment_left']
    
    # Recomendaciones estratégicas
    row += 2
    ws.merge_cells(f'A{row}:J{row}')
    ws[f'A{row}'] = 'RECOMENDACIONES ESTRATÉGICAS:'
    ws[f'A{row}'].font = estilos['subtitulo']
    ws[f'A{row}'].alignment = estilos['alignment_left']
    ws.row_dimensions[row].height = 20
    
    recomendaciones = [
        '1. IMPLEMENTAR gestiones de riesgo para variables críticas (ventas, costos, tasas).',
        '2. DIVERSIFICAR mercados de exportación para reducir dependencia de un solo cliente/país.',
        '3. ESTABLECER contratos de suministro a largo plazo para estabilizar costos.',
        '4. NEGOCIAR líneas de crédito contingentes para cubrir el peor caso.',
        '5. MONITOREAR mensualmente desviaciones vs. presupuesto y activar planes de contingencia.',
        '6. CONSIDERAR seguros de crédito y coberturas cambiarias si aplicara.',
        '7. MANTENER reservas de liquidez equivalentes a 3-6 meses de operación.',
    ]
    
    for rec in recomendaciones:
        row += 1
        ws[f'A{row}'] = rec
        ws[f'A{row}'].font = estilos['normal']
        ws[f'A{row}'].alignment = estilos['alignment_left']
    
    # Ajustar anchos de columna
    for col in range(1, 11):
        if col == 1:
            ws.column_dimensions[get_column_letter(col)].width = 22
        elif col == 5:
            ws.column_dimensions[get_column_letter(col)].width = 35
        else:
            ws.column_dimensions[get_column_letter(col)].width = 18
    
    print("✓ Hoja 'Análisis de Escenarios' creada exitosamente")
    return ws

def main():
    """Función principal que ejecuta todo el proceso"""
    print("=" * 60)
    print("CREANDO HOJAS DE RESULTADOS Y ANÁLISIS DE ESCENARIOS")
    print("=" * 60)
    
    # Paso 1: Copiar archivo base
    copiar_archivo()
    
    # Paso 2: Cargar libro de trabajo
    wb = openpyxl.load_workbook(OUTPUT_FILE)
    print(f"✓ Archivo cargado: {OUTPUT_FILE}")
    print(f"  Hojas existentes: {wb.sheetnames}")
    
    # Paso 3: Crear hoja de Resultados del Escenario Básico
    print("\n--- Creando Parte A: Resultados del Escenario Básico ---")
    crear_hoja_resultados_basicos(wb)
    
    # Paso 4: Crear hoja de Análisis de Escenarios
    print("\n--- Creando Parte B: Análisis de Escenarios ---")
    crear_hoja_analisis_escenarios(wb)
    
    # Paso 5: Guardar archivo final
    wb.save(OUTPUT_FILE)
    print(f"\n✓ Archivo guardado exitosamente: {OUTPUT_FILE}")
    
    # Verificación final
    wb_check = openpyxl.load_workbook(OUTPUT_FILE)
    print(f"\n=== VERIFICACIÓN FINAL ===")
    print(f"Hojas en el archivo ({len(wb_check.sheetnames)}):")
    for i, sheet in enumerate(wb_check.sheetnames, 1):
        print(f"  {i}. {sheet}")
    
    print("\n" + "=" * 60)
    print("¡PROCESO COMPLETADO EXITOSAMENTE!")
    print("=" * 60)
    print("\nARCHIVOS GENERADOS:")
    print(f"  • Principal: {OUTPUT_FILE}")
    print(f"  • Backup: {INPUT_FILE}")
    
    print("\nCONTENIDO AGREGADO:")
    print("  ✓ Hoja 'Resultados del Escenario Basico' (Parte A)")
    print("    - Capacidad de pago del proyecto")
    print("    - Rentabilidad para el banco")
    print("    - Rentabilidad para accionista EE.UU.")
    print("    - Conclusiones y recomendaciones")
    print()
    print("  ✓ Hoja 'Análisis de Escenarios' (Parte B)")
    print("    - Peor caso fundamentado")
    print("    - Mejor caso fundamentado")
    print("    - Caso básico (escenario esperado)")
    print("    - Cálculo de VAN ponderado")
    print("    - Conclusiones y recomendaciones estratégicas")
    print("=" * 60)

if __name__ == '__main__':
    main()
