Module M04_DetalleBonoAntiguedad
    Public Sub BonoAntiguedad(ByVal CI As String,
                                ByVal nombreCompleto As String,
                                ByVal SMN As String,
                                ByVal fechaIng As String,
                                ByVal fechaPlanilla As String)

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim filCelInc As Long
            Dim filCelFn As Long
            Dim rangoBA As Excel.Range

            ajusteHojaBA()

            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(2, 0).Row
            filCelInc = n + 1
            filCelFn = n + 15

            rangoBA = .Range(.Cells(n, 1), .Cells(filCelFn, 13))
            rangoBA.Select()

            contenedorBA()

            '#######################################################
            '#############    TITULO DE REPORTE
            '#######################################################

            .Cells(n, 1).Value = "CALCULO BONO DE ANTIGUEDAD"
            .Cells(filCelFn, 1).Value = "Conforme al articulo 60 del D.S. 21060 del 29 de agosto de 1986."

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


            '#######################################################
            '#############    TABLA DE ANTIGUEDAD
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(1, 1), .Cells(filCelInc, 1).Offset(2, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "AÑOS"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(1, 3), .Cells(filCelInc, 1).Offset(2, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "PORCENTAJE"
            End With

            '#############    DE DOS A CUATROA AÑOS 5%
            With .Range(.Cells(filCelInc, 1).Offset(3, 1), .Cells(filCelInc, 1).Offset(3, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "2"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(3, 2), .Cells(filCelInc, 1).Offset(3, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "4"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(3, 3), .Cells(filCelInc, 1).Offset(3, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "5%"
            End With

            '#############    DE 5 A 7 AÑOS 11%
            With .Range(.Cells(filCelInc, 1).Offset(4, 1), .Cells(filCelInc, 1).Offset(4, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "5"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 2), .Cells(filCelInc, 1).Offset(4, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "7"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 3), .Cells(filCelInc, 1).Offset(4, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "11%"
            End With

            '#############    DE 8 A 10 AÑOS 18%
            With .Range(.Cells(filCelInc, 1).Offset(5, 1), .Cells(filCelInc, 1).Offset(5, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "8"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(5, 2), .Cells(filCelInc, 1).Offset(5, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "10"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(5, 3), .Cells(filCelInc, 1).Offset(5, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "18%"
            End With

            '#############    DE 11 A 14 AÑOS 26%
            With .Range(.Cells(filCelInc, 1).Offset(6, 1), .Cells(filCelInc, 1).Offset(6, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "11"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 2), .Cells(filCelInc, 1).Offset(6, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "14"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 3), .Cells(filCelInc, 1).Offset(6, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "26%"
            End With

            '#############    DE 15 A 19 AÑOS 34%
            With .Range(.Cells(filCelInc, 1).Offset(7, 1), .Cells(filCelInc, 1).Offset(7, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "15"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 2), .Cells(filCelInc, 1).Offset(7, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "34"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 3), .Cells(filCelInc, 1).Offset(7, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "34%"
            End With

            '#############    DE 20 A 24 AÑOS 42%
            With .Range(.Cells(filCelInc, 1).Offset(7, 1), .Cells(filCelInc, 1).Offset(7, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "20"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 2), .Cells(filCelInc, 1).Offset(7, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "24"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(7, 3), .Cells(filCelInc, 1).Offset(7, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "42%"
            End With

            '#############    DE 25 A MAS AÑOS 50%
            With .Range(.Cells(filCelInc, 1).Offset(8, 1), .Cells(filCelInc, 1).Offset(8, 1))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "25"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 2), .Cells(filCelInc, 1).Offset(8, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "A MAS"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 3), .Cells(filCelInc, 1).Offset(8, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "50%"
            End With

            .Range(.Cells(filCelInc, 1).Offset(1, 1), .Cells(filCelInc, 1).Offset(2, 2)).CurrentRegion.Select
            formatoTablas()

            '#######################################################
            '#############    TABLA SMN - 3SMN
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(3, 5), .Cells(filCelInc, 1).Offset(3, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "S.M.N."
                .Offset(0, 1).Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 5), .Cells(filCelInc, 1).Offset(4, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "3 S.M.N."
                .Offset(0, 1).FormulaR1C1 = "=3*R[-1]C"
                .Offset(0, 1).Font.Bold = True
                .Offset(0, 1).NumberFormat = "#,##0.00"
                .Offset(0, 1).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            .Range(.Cells(filCelInc, 1).Offset(3, 5), .Cells(filCelInc, 1).Offset(3, 5)).CurrentRegion.Select
            formatoTablas()

            '#######################################################
            '#############    TABLA CALCULO ANTIGUEDAD
            '#######################################################

            With .Range(.Cells(filCelInc, 1).Offset(1, 8), .Cells(filCelInc, 1).Offset(1, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "CALCULO ANTIGUEDAD"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 8), .Cells(filCelInc, 1).Offset(2, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "FECHA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 9), .Cells(filCelInc, 1).Offset(2, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "AÑOS"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 10), .Cells(filCelInc, 1).Offset(2, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "MESES"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 11), .Cells(filCelInc, 1).Offset(2, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "DIAS"
            End With

            With .Range(.Cells(filCelInc, 1).Offset(3, 8), .Cells(filCelInc, 1).Offset(3, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "-"
            End With

            With .Range(.Cells(filCelInc, 1).Offset(4, 8), .Cells(filCelInc, 1).Offset(4, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "Ingreso"
            End With

            With .Range(.Cells(filCelInc, 1).Offset(5, 8), .Cells(filCelInc, 1).Offset(5, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "Resultado"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(5, 9), .Cells(filCelInc, 1).Offset(5, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=R[-2]C-R[-1]C"
                .NumberFormat = "#,##0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(5, 10), .Cells(filCelInc, 1).Offset(5, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=R[-2]C-R[-1]C"
                .NumberFormat = "#,##0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(5, 11), .Cells(filCelInc, 1).Offset(5, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=R[-2]C-R[-1]C"
                .NumberFormat = "#,##0"
            End With

            .Range(.Cells(filCelInc, 1).Offset(1, 8), .Cells(filCelInc, 1).Offset(1, 11)).CurrentRegion.Select
            formatoTablas()

            '#######################################################
            '#############    TABLA CALCULO BONO DEL TRABAJADOR
            '#######################################################

            'TITULOS TABLA
            With .Range(.Cells(filCelInc, 1).Offset(11, 2), .Cells(filCelInc, 1).Offset(11, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "C.I."
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 3), .Cells(filCelInc, 1).Offset(11, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "EMPLEADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 5), .Cells(filCelInc, 1).Offset(11, 6))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "FECHA INGRESO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 7), .Cells(filCelInc, 1).Offset(11, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "FECHA PLANILLA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 8), .Cells(filCelInc, 1).Offset(11, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "AÑOS ANT."
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 9), .Cells(filCelInc, 1).Offset(11, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "%"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(11, 10), .Cells(filCelInc, 1).Offset(11, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "BONO"
            End With
            'VALORES TABLA

            With .Range(.Cells(filCelInc, 1).Offset(12, 2), .Cells(filCelInc, 1).Offset(12, 2))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 3), .Cells(filCelInc, 1).Offset(12, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 5), .Cells(filCelInc, 1).Offset(12, 6))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 7), .Cells(filCelInc, 1).Offset(12, 7))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 8), .Cells(filCelInc, 1).Offset(12, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 9), .Cells(filCelInc, 1).Offset(12, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 10), .Cells(filCelInc, 1).Offset(12, 10))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "0"
            End With

            .Range(.Cells(filCelInc, 1).Offset(11, 2), .Cells(filCelInc, 1).Offset(11, 2)).CurrentRegion.Select
            formatoTablas()


            calculoBA(filCelInc,
                CI,
                nombreCompleto,
                SMN,
                fechaIng,
                fechaPlanilla)

        End With

    End Sub
    Sub ajusteHojaBA()

        With Globals.ThisAddIn.Application

            .Columns("A:A").ColumnWidth = 6
            .Columns("M:M").ColumnWidth = 6
            .Columns("D:D").ColumnWidth = 20
            .Columns("H:H").ColumnWidth = 20
            .Columns("I:I").ColumnWidth = 20

        End With

    End Sub
    Sub contenedorBA()

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
    Sub formatoTablas()

        With Globals.ThisAddIn.Application

            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With
            With .Selection.Borders(Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal)
                .LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                .ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .TintAndShade = 0
                .Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
            End With

        End With

    End Sub

    Sub calculoBA(ByVal filCelInc As Long,
              ByVal CI As String,
              ByVal nombreCompleto As String,
              ByVal SMN As String,
              ByVal fechaIng As String,
              ByVal fechaPlanilla As String)

        With Globals.ThisAddIn.Application

            'CALUCULO DE ANTIGUEDAD
            'FECHA DE LA PLANILLA
            With .Range(.Cells(filCelInc, 1).Offset(3, 9), .Cells(filCelInc, 1).Offset(3, 9))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[9]C[-2],""yyyy"")" &
                            "-IF(TEXT(R[9]C[-2],""m"")<TEXT(R[9]C[-4],""m""),1,0)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(3, 10), .Cells(filCelInc, 1).Offset(3, 10))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[9]C[-3],""m"")" &
                                "-IF(TEXT(R[9]C[-3],""d"")<TEXT(R[9]C[-5],""d""),1,0)" &
                                "+IF(TEXT(R[9]C[-3],""m"")<TEXT(R[9]C[-5],""m""),12,0)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(3, 11), .Cells(filCelInc, 1).Offset(3, 11))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[9]C[-4],""d"")" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""1"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""2"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),28,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""3"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""4"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""5"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""6"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""7"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""8"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""9"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""10"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""11"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)<>0,TEXT(R[9]C[-4],""m"")=""12"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""1"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""2"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),29,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""3"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""4"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""5"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""6"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""7"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""8"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""9"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""10"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""11"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),30,0)" &
                                        "+ IF(AND(MOD(TEXT(R[9]C[-4],""m""),4)=0,TEXT(R[9]C[-4],""m"")=""12"",TEXT(R[9]C[-4],""d"")<TEXT(R[9]C[-6],""d"")),31,0)"
            End With

            'FECHA INGRESO
            With .Range(.Cells(filCelInc, 1).Offset(4, 9), .Cells(filCelInc, 1).Offset(4, 9))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[8]C[-4],""yyyy"")"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 10), .Cells(filCelInc, 1).Offset(4, 10))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[8]C[-5],""m"")"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 11), .Cells(filCelInc, 1).Offset(4, 11))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=TEXT(R[8]C[-6],""d"")"
            End With

            ' CALCULO DE BONO DE ANTIGUESDAD

            With .Range(.Cells(filCelInc, 1).Offset(12, 2), .Cells(filCelInc, 1).Offset(12, 2))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = CI
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 3), .Cells(filCelInc, 1).Offset(12, 4))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = nombreCompleto
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 5), .Cells(filCelInc, 1).Offset(12, 6))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = fechaIng
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 7), .Cells(filCelInc, 1).Offset(12, 7))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = fechaPlanilla
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 8), .Cells(filCelInc, 1).Offset(12, 8))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=DATEDIF(RC[-3],RC[-1],""Y"")"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 9), .Cells(filCelInc, 1).Offset(12, 9))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .NumberFormat = "0%"
                .FormulaR1C1 = "=IF(AND(RC[-1]>=R[-9]C[-8],RC[-1]<=R[-9]C[-7]),R[-9]C[-6],IF(AND(RC[-1]>=R[-8]C[-8],RC[-1]<=R[-8]C[-7]),R[-8]C[-6],IF(AND(RC[-1]>=R[-7]C[-8],RC[-1]<=R[-7]C[-7]),R[-7]C[-6],IF(AND(RC[-1]>=R[-6]C[-8],RC[-1]<=R[-6]C[-7]),R[-6]C[-6],IF(AND(RC[-1]>=R[-5]C[-8],RC[-1]<=R[-5]C[-7]),R[-5]C[-6],IF(AND(RC[-1]>=R[-4]C[-8]),R[-4]C[-6],0))))))"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 10), .Cells(filCelInc, 1).Offset(12, 10))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .FormulaR1C1 = "=RC[-1]*R[-8]C[-4]"
                .NumberFormat = "#,##0.00"
            End With

            ' SALARIO MINIMO NACIONAL
            With .Range(.Cells(filCelInc, 1).Offset(3, 6), .Cells(filCelInc, 1).Offset(3, 6))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = SMN
                .NumberFormat = "#,##0.00"
            End With

        End With
    End Sub

    Public Function resultadoBoAnt(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultadoBoAnt = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(12, 10).Address
        End With
    End Function

End Module
