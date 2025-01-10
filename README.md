# Macro para aplicar un Heat Map por fila en un PivotTable de Excel

Este script VBA aplica un **mapa de calor** por fila a un PivotTable en Excel, formateando uniformemente todas las filas (incluidos los totales) mediante una escala de colores suave.

## Características

- **Formato Condicional:** Aplica un mapa de calor por fila entre:
  - `"PRI"` y la columna anterior a `"PRI Total"`.
  - `"PRI1H"` y la columna anterior a `"PRI1H Total"`.
- **Escala de colores:**
  - **Verde:** Valores altos.
  - **Amarillo:** Valores medios.
  - **Rojo:** Valores bajos.
- Compatible con todas las filas, incluidas las de totales.

---

## Código VBA

```vba
Sub CreateRowHeatMapForPivotTable()
    Dim ws As Worksheet
    Dim rowStart As Long, rowEnd As Long
    Dim priStartCol As Long, priEndCol As Long
    Dim pri1hStartCol As Long, pri1hEndCol As Long
    Dim i As Long, rng As Range

    ' Configuración de hoja y rangos
    Set ws = ThisWorkbook.Sheets(1)
    rowStart = 2: rowEnd = 200
    priStartCol = ws.Range("B1").Column: priEndCol = ws.Range("L1").Column
    pri1hStartCol = ws.Range("N1").Column: pri1hEndCol = ws.Range("S1").Column

    ' Recorre las filas y aplica el mapa de calor
    For i = rowStart To rowEnd
        Set rng = ws.Range(ws.Cells(i, priStartCol), ws.Cells(i, priEndCol))
        ApplyHeatMap rng
        Set rng = ws.Range(ws.Cells(i, pri1hStartCol), ws.Cells(i, pri1hEndCol))
        ApplyHeatMap rng
    Next i
End Sub

Sub ApplyHeatMap(rng As Range)
    rng.FormatConditions.Delete
    With rng.FormatConditions.AddColorScale(3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 153, 153)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 204)
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = RGB(204, 255, 204)
    End With
End Sub
```


## Ejecución
1.	Abrir el Editor VBA (Alt + F11).
2.	Insertar Módulo: Ve a Insertar > Módulo y pega el código.
3.	Cerrar el Editor (Alt + Q).
4.	Ejecutar el Macro:
- Presiona Alt + F8, selecciona CreateRowHeatMapForPivotTable y haz clic en Ejecutar.