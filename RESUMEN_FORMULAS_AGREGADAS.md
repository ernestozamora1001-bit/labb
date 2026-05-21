# Resumen de Fórmulas Agregadas - Caso Hamacas

## Archivo: Medrano_Zamora_COMPLETO.xlsx
## Fecha de Actualización: Mayo 2026
## Scripts Disponibles: apply_all_formulas.py (script maestro)

---

## Cuadros Completados

### CUADRO 4: Requerimiento de Materiales y Mano de Obra (Filas 100-105)

Se agregaron fórmulas para referenciar valores desde CUADRO 7:

| Celda | Fórmula Agregada | Descripción |
|-------|------------------|-------------|
| C100 | =D179 | Repuestos y Mantenimiento (Fase 1) |
| D100 | =E179 | Repuestos y Mantenimiento (Fase 2) |
| E100 | =F179 | Repuestos y Mantenimiento (Fases 3-5) |
| C101 | =D179 | Gastos Generales (Fase 1) |
| D101 | =E179 | Gastos Generales (Fase 2) |
| E101 | =F179 | Gastos Generales (Fases 3-5) |
| C103 | =D184 | Administración General (Fase 1) |
| D103 | =E184 | Administración General (Fase 2) |
| E103 | =F184 | Administración General (Fases 3-5) |
| C104 | =D187 | Gasto de Ventas (Fase 1) |
| D104 | =E187 | Gasto de Ventas (Fase 2) |
| E104 | =F187 | Gasto de Ventas (Fases 3-5) |
| C105 | =D191 | Gasto de Distribución (Fase 1) |
| D105 | =E191 | Gasto de Distribución (Fase 2) |
| E105 | =F191 | Gasto de Distribución (Fases 3-5) |

**Nota:** Los valores en CUADRO 7 fila 191 ya existen:
- D191: 60,000
- E191: 80,000  
- F191: 100,000

---

### CUADRO 5: Requerimientos de Capital de Trabajo (Filas 111-134)

#### ACTIVO CIRCULANTE:

| Celda | Fórmula | Descripción |
|-------|---------|-------------|
| D111 | =($E$88*100/360)*C111 | Cuentas por Cobrar |
| D114 | =(E95*2/360)*C114 | Materiales Locales |
| D115 | =(E96*2/360)*C115 | Materiales Importados |
| D116 | =(E100*2/360)*C116 | Repuestos |
| D117 | =((E95+E96+E97+E98)/2/360)*C117 | Productos en Proceso |
| D118 | =((E95+E96+E97+E98)*2/360)*C118 | Producto Terminado |
| D120 | =D111+D114+D115+D116+D117+D118 | Requerimiento de Liquidez |
| D121 | =SUM(D111:D118) | **Total Activo Circulante** |

#### PASIVO CIRCULANTE:

| Celda | Fórmula | Descripción |
|-------|---------|-------------|
| D126 | =(E97*2/360)*C126 | Mano de Obra |
| D127 | =(E98*2/360)*C127 | Servicios |
| D129 | =((E184+E187+E191)/2/360)*C129 | Gastos Admin, Ventas y Distrib. |
| D131 | =I95*2/360*C131 | Gastos Financieros |
| D132 | =SUM(D124:D131) | **Total Pasivo Circulante** |

#### CAPITAL DE TRABAJO:

| Celda | Fórmula | Descripción |
|-------|---------|-------------|
| D134 | =D121-D132 | **Capital de Trabajo Neto Requerido** |

---

### CUADRO 6: Costos Fijos y Variables (Filas 153-154)

Se agregaron fórmulas de totales:

| Celda | Fórmula | Descripción |
|-------|---------|-------------|
| B153 | TOTAL COSTOS FIJOS | Etiqueta |
| C153 | =SUMIF(C142:C152,"X",C142:C152) | Total Costos Fijos |
| D153 | =SUMIF(D142:D152,"X",D142:D152) | Total Costos Variables |
| B154 | COSTO TOTAL | Etiqueta |
| C154 | =C153 | Costo Total Fijo |
| D154 | =D153 | Costo Total Variable |

---

## Fórmulas Existentes Preservadas

Todas las fórmulas existentes en el archivo fueron preservadas, incluyendo:
- CUADRO 1: Todas las fórmulas de totales y referencias
- CUADRO 2: Fórmulas de inversión y depreciación
- CUADRO 3: Fórmulas del programa de producción
- CUADRO 7: Fórmulas de costos de operación

---

## Cálculos Utilizados

### Capital de Trabajo:
```
Capital de Trabajo = (Costo Anual / 360) × Días de Inventario
```

Donde:
- Costo Anual = Valor de la fase máxima (columna E) × 2 semestres
- Días = Días promedio de rotación según cada concepto

### Activo vs Pasivo Circulante:
```
Capital de Trabajo Neto = Activo Circulante - Pasivo Circulante
```

---

## Referencias entre Cuadros

- CUADRO 4 → CUADRO 7: Referencias a costos de operación (filas 179, 184, 187, 191)
- CUADRO 5 → CUADRO 4: Referencias a materiales y mano de obra (filas 95-98, 100)
- CUADRO 5 → CUADRO 7: Referencias a gastos administrativos (filas 184, 187, 191)
- CUADRO 6: Tabla de clasificación con totales SUMIF

---

## Estado Final

✅ **CUADRO 1**: Completo (fórmulas existentes preservadas)
✅ **CUADRO 2**: Completo (fórmulas existentes preservadas)
✅ **CUADRO 3**: Completo (fórmulas existentes preservadas)
✅ **CUADRO 4**: Completado con 15 fórmulas de referencia
✅ **CUADRO 5**: Completado con 14 fórmulas de cálculo
✅ **CUADRO 6**: Completado con 4 fórmulas de totales
✅ **CUADRO 7**: Completo (fórmulas existentes)

**Total de fórmulas agregadas: 33**

---

## Archivos Generados

| Archivo | Descripción |
|---------|-------------|
| `Medrano_Zamora_COMPLETO.xlsx` | Archivo final con todas las fórmulas aplicadas |
| `Medrano_Zamora_BACKUP.xlsx` | Backup de seguridad antes de aplicar cambios |
| `apply_all_formulas.py` | Script maestro para aplicar todas las fórmulas |

---

## Cómo Usar

1. **Script Maestro**: Ejecutar `python3 apply_all_formulas.py` para aplicar todas las fórmulas
2. **Archivo Excel**: Abrir `Medrano_Zamora_COMPLETO.xlsx` en Excel
3. Las fórmulas se calcularán automáticamente al abrir el archivo

---

**Nota:** Las fórmulas utilizan la sintaxis de Excel estándar (formato con "=" al inicio) y están diseñadas para calcular automáticamente cuando se abra el archivo en Excel.

**Última actualización:** Mayo 2026
