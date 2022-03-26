Module M03_DetalleANS
    Public Sub APN(ByVal totalGanado As String, Optional ByVal nombreCompleto As String = "XXXX", Optional ByVal carnet As String = "XXXXX")

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim filCelInc As Long
            Dim filCelFn As Long
            Dim rangoAPN As Excel.Range

            ajusteHoja()

            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(2, 0).Row
            filCelInc = n + 1
            filCelFn = n + 29

            rangoAPN = .Range(.Cells(n, 1), .Cells(filCelFn, 13))
            rangoAPN.Select()

            contenedor()
            calculoANS(n, filCelFn, filCelInc, totalGanado, nombreCompleto, carnet)
        End With

    End Sub

    Public Sub ajusteHoja()

        With Globals.ThisAddIn.Application
            Dim columnas(5) As String
            Dim i As Integer

            columnas(0) = "A"
            columnas(1) = "C"
            columnas(2) = "E"
            columnas(3) = "I"
            columnas(4) = "K"
            columnas(5) = "M"

            For i = 0 To 5
                .Columns(columnas(i) & ":" & columnas(i)).ColumnWidth = 3
            Next

            .Columns("G:G").ColumnWidth = 9
        End With
    End Sub

    Public Sub contenedor()
        With Globals.ThisAddIn.Application
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Core.XlBorderWeight.xlThick
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Core.XlBorderWeight.xlThick
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Core.XlBorderWeight.xlThick
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDouble
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Core.XlBorderWeight.xlThick
            End With
        End With
    End Sub

    Public Sub calculoANS(ByVal n As Long, ByVal filCelFn As Long, ByVal filCelInc As Long, ByVal totalGanado As String, ByVal nombreCompleto As String, ByVal carnet As String)
        With Globals.ThisAddIn.Application
            '#######################################################
            '#############    TITULO DE REPORTE
            '#######################################################

            .Cells(n, 1).Value = "REPORTE DE APORTE NACIONAL SOLIDARIO"
            .Cells(filCelFn, 1).Value = "CONFORME A LEY 065. APORTE NACIONAL SOLIDARIO."

            With .Range(.Cells(n, 1), .Cells(n, 13))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 16
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With
            With .Range(.Cells(filCelFn, 1), .Cells(filCelFn, 13))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With

            .Cells(filCelInc, 1).Offset(0, 1).Value = "NOM. COMP.:"
            .Cells(filCelInc, 1).Offset(0, 1).Font.Bold = True
            .Cells(filCelInc, 1).Offset(0, 3).Value = nombreCompleto

            .Cells(filCelInc, 1).Offset(1, 1).Value = "C.I.:"
            .Cells(filCelInc, 1).Offset(1, 1).Font.Bold = True
            .Cells(filCelInc, 1).Offset(1, 3).Value = carnet

            .Cells(filCelInc, 1).Offset(0, 7).Value = "TOTAL GANADO:"
            .Cells(filCelInc, 1).Offset(0, 7).Font.Bold = True
            .Cells(filCelInc, 1).Offset(0, 9).Value = totalGanado
            .Cells(filCelInc, 1).Offset(0, 9).NumberFormat = "#,##0.00"


            '#############    APORTE SOLIDARIO 1%

            With .Range(.Cells(filCelInc, 1).Offset(3, 1), .Cells(filCelInc, 1).Offset(3, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "APORTE SOLIDARIO (1%)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 1), .Cells(filCelInc, 1).Offset(4, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "((Total Ganado - Bs. 13000)*1%)"
            End With

            '#############    APORTE SOLIDARIO 5%

            With .Range(.Cells(filCelInc, 1).Offset(16, 1), .Cells(filCelInc, 1).Offset(16, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "APORTE SOLIDARIO (5%)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(17, 1), .Cells(filCelInc, 1).Offset(17, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "((Total Ganado - Bs. 25000)*5%)"
            End With

            '#############    APORTE SOLIDARIOS 10%

            With .Range(.Cells(filCelInc, 1).Offset(3, 7), .Cells(filCelInc, 1).Offset(3, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "APORTE SOLIDARIO (10%)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 7), .Cells(filCelInc, 1).Offset(4, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "((Total Ganado - Bs. 35000)*10%)"
            End With

            '#############    TOTAL APORTE NACIONAL SOLIDARIO

            With .Range(.Cells(filCelInc, 1).Offset(16, 7), .Cells(filCelInc, 1).Offset(16, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "TOTAL APORTE SOLIDARIO NACIONAL"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(17, 7), .Cells(filCelInc, 1).Offset(17, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "A.N.S = LA SUMA DE LOS TRES RESULTADOS"
            End With

            '#######################################################
            '#############    TOTAL GANADO
            '#######################################################

            '#############    APORTE SOLIDARIO 1%

            With .Range(.Cells(filCelInc, 1).Offset(6, 1), .Cells(filCelInc, 1).Offset(6, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "T. GANDADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 1), .Cells(filCelInc, 1).Offset(7, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Value = totalGanado
            End With

            '#############    APORTE SOLIDARIO 5%

            With .Range(.Cells(filCelInc, 1).Offset(19, 1), .Cells(filCelInc, 1).Offset(19, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "T. GANDADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(20, 1), .Cells(filCelInc, 1).Offset(20, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Value = totalGanado
            End With

            '#############    APORTE SOLIDARIO 10%
            With .Range(.Cells(filCelInc, 1).Offset(6, 7), .Cells(filCelInc, 1).Offset(6, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "T. GANDADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 7), .Cells(filCelInc, 1).Offset(7, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Value = totalGanado
            End With

            '#######################################################
            '#############    13000, 1%
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(6, 2), .Cells(filCelInc, 1).Offset(7, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "-"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 3), .Cells(filCelInc, 1).Offset(7, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = 13000
                .NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 4), .Cells(filCelInc, 1).Offset(7, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 5), .Cells(filCelInc, 1).Offset(7, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "1%"
            End With

            '#######################################################
            '#############    25000, 5%
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(19, 2), .Cells(filCelInc, 1).Offset(20, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "-"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(19, 3), .Cells(filCelInc, 1).Offset(20, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = 25000
                .NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(19, 4), .Cells(filCelInc, 1).Offset(20, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(19, 5), .Cells(filCelInc, 1).Offset(20, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "5%"
            End With
            '#######################################################
            '#############    35000, 10%
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(6, 8), .Cells(filCelInc, 1).Offset(7, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "-"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 9), .Cells(filCelInc, 1).Offset(7, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = 35000
                .NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 10), .Cells(filCelInc, 1).Offset(7, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 11), .Cells(filCelInc, 1).Offset(7, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "10%"
            End With

            '#######################################################
            '#############    DIFERENCIA
            '#######################################################

            '#############    APORTE SOLIDARIO 1%

            With .Range(.Cells(filCelInc, 1).Offset(9, 1), .Cells(filCelInc, 1).Offset(9, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "DIFERENCIA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 1), .Cells(filCelInc, 1).Offset(10, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C-R[-4]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(9, 4), .Cells(filCelInc, 1).Offset(10, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(9, 5), .Cells(filCelInc, 1).Offset(10, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "1%"
            End With

            '#############    APORTE SOLIDARIO 5%

            With .Range(.Cells(filCelInc, 1).Offset(22, 1), .Cells(filCelInc, 1).Offset(22, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "DIFERENCIA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(23, 1), .Cells(filCelInc, 1).Offset(23, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C-R[-4]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(22, 4), .Cells(filCelInc, 1).Offset(23, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(22, 5), .Cells(filCelInc, 1).Offset(23, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "5%"
            End With

            '#############    APORTE SOLIDARIO 10%

            With .Range(.Cells(filCelInc, 1).Offset(9, 7), .Cells(filCelInc, 1).Offset(9, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "DIFERENCIA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 7), .Cells(filCelInc, 1).Offset(10, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C-R[-4]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(9, 10), .Cells(filCelInc, 1).Offset(10, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(9, 11), .Cells(filCelInc, 1).Offset(10, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "10%"
            End With


            '#######################################################
            '#############    RESULTADO
            '#######################################################

            '#############    APORTE SOLIDARIO 1%

            With .Range(.Cells(filCelInc, 1).Offset(12, 1), .Cells(filCelInc, 1).Offset(12, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(13, 1), .Cells(filCelInc, 1).Offset(13, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C*R[-4]C[4]"
                .NumberFormat = "#,##0.00"
            End With

            '#############    APORTE SOLIDARIO 5%

            With .Range(.Cells(filCelInc, 1).Offset(25, 1), .Cells(filCelInc, 1).Offset(25, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(26, 1), .Cells(filCelInc, 1).Offset(26, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C*R[-4]C[4]"
                .NumberFormat = "#,##0.00"
            End With

            '#############    APORTE SOLIDARIO 10%

            With .Range(.Cells(filCelInc, 1).Offset(12, 7), .Cells(filCelInc, 1).Offset(12, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(13, 7), .Cells(filCelInc, 1).Offset(13, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C*R[-4]C[4]"
                .NumberFormat = "#,##0.00"
            End With

            '#######################################################
            '#############    RESULTADOS
            '#######################################################
            '#############    APORTE SOLIDARIO 1%
            With .Range(.Cells(filCelInc, 1).Offset(19, 7), .Cells(filCelInc, 1).Offset(19, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO 1"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(20, 7), .Cells(filCelInc, 1).Offset(20, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=IF(R[-7]C[-6]<0,0,R[-7]C[-6])"
                .NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(19, 8), .Cells(filCelInc, 1).Offset(20, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "+"
            End With
            '#############    APORTE SOLIDARIO 5%
            With .Range(.Cells(filCelInc, 1).Offset(19, 9), .Cells(filCelInc, 1).Offset(19, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO 2"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(20, 9), .Cells(filCelInc, 1).Offset(20, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=IF(R[6]C[-8]<0,0,R[6]C[-8])"
                .NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(19, 10), .Cells(filCelInc, 1).Offset(20, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Value = "+"
            End With
            '#############    APORTE SOLIDARIO 10%
            With .Range(.Cells(filCelInc, 1).Offset(19, 11), .Cells(filCelInc, 1).Offset(19, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "RESULTADO 3"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(20, 11), .Cells(filCelInc, 1).Offset(20, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=IF(R[-7]C[-4]<0,0,R[-7]C[-4])"
                .NumberFormat = "#,##0.00"
            End With

            '#############   TOTAL APORTE NACIONAL SOLIDARIO
            With .Range(.Cells(filCelInc, 1).Offset(22, 7), .Cells(filCelInc, 1).Offset(22, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
                .Value = "TOTAL A.N.S DEL TRABAJADOR"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(23, 7), .Cells(filCelInc, 1).Offset(23, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .FormulaR1C1 = "=R[-3]C+R[-3]C[2]+R[-3]C[4]"
                .NumberFormat = "#,##0.00"
            End With
        End With
    End Sub

End Module
