Module M06_DetalleHorasExtraordinarias
    Public Sub detallaHorasExtraordinarias(ByVal categoriaTrabNoct As String,
                            ByVal CI As String,
                            ByVal nombreCompleto As String,
                            ByVal haberBasico As String,
                            ByVal sexo As String,
                            ByVal horasExtraordinarias As String,
                            ByVal horasNocturnas As String)
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim filCelInc As Long
            Dim filCelFn As Long
            Dim rangoHrExt As Excel.Range

            ajusteHojaHrExt()

            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(2, 0).Row
            filCelInc = n + 1
            filCelFn = n + 20

            rangoHrExt = .Range(.Cells(n, 1), .Cells(filCelFn, 11))
            rangoHrExt.Select()

            contenedorHrExt()

            '##############################################################
            '#############    TITULO DE REPORTE HORAS EXTRAORDINARIAS
            '##############################################################

            .Cells(n, 1).Value = "TRABAJO EXTRAORDINARIO Y NOCTURNO"
            .Cells(filCelFn, 1).Value = "CONFORME AL ARTICULO 55. LEY GENERAL DEL TRABAJO."

            With .Range(.Cells(n, 1), .Cells(n, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 20
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With
            With .Range(.Cells(filCelFn, 1), .Cells(filCelFn, 11))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 18
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With

            '##############################################################
            '#############    DATOS DEL INFORME
            '##############################################################
            With .Range(.Cells(filCelInc, 1).Offset(0, 1), .Cells(filCelInc, 1).Offset(0, 1))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "EMPLEADO:"
                .Offset(0, 2).Value = nombreCompleto
            End With
            With .Range(.Cells(filCelInc, 1).Offset(1, 1), .Cells(filCelInc, 1).Offset(1, 1))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "C.I.:"
                .Offset(0, 2).Value = CI
                .Offset(0, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 1), .Cells(filCelInc, 1).Offset(2, 1))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "SEXO:"
                .Offset(0, 2).Value = sexo
            End With
            With .Range(.Cells(filCelInc, 1).Offset(0, 6), .Cells(filCelInc, 1).Offset(0, 6))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "HABER BASICO:"
                .Offset(0, 2).Value = haberBasico
            End With
            With .Range(.Cells(filCelInc, 1).Offset(1, 6), .Cells(filCelInc, 1).Offset(1, 6))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "CATEGORIA DE TRABAJO NOCTURNO:"
                .Offset(1, 0).Value = categoriaTrabNoct
            End With

            '###################################################################
            '#############   TITULOS HORA EXTRAS, RECARGO NOCTURNO
            '###################################################################

            With .Range(.Cells(filCelInc, 1).Offset(4, 1), .Cells(filCelInc, 1).Offset(4, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "CALCULO HORA EXTRA"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 6), .Cells(filCelInc, 1).Offset(4, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "CALCULO DE RECARGO NOCTURNO"
            End With
            '###################################################################
            '#############   ETIQUETAS HORA EXTRAS
            '###################################################################

            'FILA DIAS
            With .Range(.Cells(filCelInc, 1).Offset(6, 1), .Cells(filCelInc, 1).Offset(6, 1))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-6]C[7]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 2), .Cells(filCelInc, 1).Offset(6, 2))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "/"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 3), .Cells(filCelInc, 1).Offset(6, 3))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "30"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 4), .Cells(filCelInc, 1).Offset(6, 4))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "DIAS"
            End With

            'FILA HORAS

            With .Range(.Cells(filCelInc, 1).Offset(8, 1), .Cells(filCelInc, 1).Offset(8, 1))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C/R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 2), .Cells(filCelInc, 1).Offset(8, 2))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "/"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 3), .Cells(filCelInc, 1).Offset(8, 3))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "8"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 4), .Cells(filCelInc, 1).Offset(8, 4))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "HORAS"
            End With

            'FILA DOBLE

            With .Range(.Cells(filCelInc, 1).Offset(10, 1), .Cells(filCelInc, 1).Offset(10, 1))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C/R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 2), .Cells(filCelInc, 1).Offset(10, 2))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 3), .Cells(filCelInc, 1).Offset(10, 3))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "2"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 4), .Cells(filCelInc, 1).Offset(10, 4))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "DOBLE"
            End With

            'FILA N° H. EXT.
            With .Range(.Cells(filCelInc, 1).Offset(12, 1), .Cells(filCelInc, 1).Offset(12, 1))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C*R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 2), .Cells(filCelInc, 1).Offset(12, 2))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 3), .Cells(filCelInc, 1).Offset(12, 3))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .NumberFormat = "#,##0.00"
                .Value = horasExtraordinarias
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 4), .Cells(filCelInc, 1).Offset(12, 4))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "N° H. EXT."
            End With

            ' TOTAL HORAS EXTRAS

            With .Range(.Cells(filCelInc, 1).Offset(14, 1), .Cells(filCelInc, 1).Offset(14, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "TOTAL POR HORAS EXTRAS"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(15, 1), .Cells(filCelInc, 1).Offset(15, 4))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-3]C*R[-3]C[2]"
            End With

            '###################################################################
            '#############   ETIQUETAS RECARGO NOCTURNO
            '###################################################################

            'FILA DIAS
            With .Range(.Cells(filCelInc, 1).Offset(6, 6), .Cells(filCelInc, 1).Offset(6, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-6]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 7), .Cells(filCelInc, 1).Offset(6, 7))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "/"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 8), .Cells(filCelInc, 1).Offset(6, 8))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "30"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 9), .Cells(filCelInc, 1).Offset(6, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "DIAS"
            End With

            'FILA HORAS

            With .Range(.Cells(filCelInc, 1).Offset(8, 6), .Cells(filCelInc, 1).Offset(8, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C/R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 7), .Cells(filCelInc, 1).Offset(8, 7))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "/"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 8), .Cells(filCelInc, 1).Offset(8, 8))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "8"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 9), .Cells(filCelInc, 1).Offset(8, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "HORAS"
            End With

            'FILA RECARGO

            With .Range(.Cells(filCelInc, 1).Offset(10, 6), .Cells(filCelInc, 1).Offset(10, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C/R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 7), .Cells(filCelInc, 1).Offset(10, 7))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 8), .Cells(filCelInc, 1).Offset(10, 8))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .FormulaR1C1 = "=IF(R[-8]C[-2]=""HOMBRE ADMINISTRATIVO"",""0,25"",""0"")" &
                        "+IF(R[-8]C[-2]=""HOMBRE OBRERO"",""0,30"",""0"")" &
                        "+IF(R[-8]C[-2]=""MUJERES O MENOR 18 AÑOS"",""0,40"",""0"")" &
                        "+IF(R[-8]C[-2]=""TRABAJO DE ALTO RIESGO"",""0,50"",""0"")"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(10, 9), .Cells(filCelInc, 1).Offset(10, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "RECARGO"
            End With

            'FILA N° H. NOCTURNAS
            With .Range(.Cells(filCelInc, 1).Offset(12, 6), .Cells(filCelInc, 1).Offset(12, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C*R[-2]C[2]"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 7), .Cells(filCelInc, 1).Offset(12, 7))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "X"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 8), .Cells(filCelInc, 1).Offset(12, 8))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .NumberFormat = "#,##0.00"
                .Value = horasNocturnas
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 9), .Cells(filCelInc, 1).Offset(12, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "N° H. NOCT."
            End With

            ' TOTAL HORAS NOCTURNAS

            With .Range(.Cells(filCelInc, 1).Offset(14, 6), .Cells(filCelInc, 1).Offset(14, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "TOTAL RECARGO NOCTURNO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(15, 6), .Cells(filCelInc, 1).Offset(15, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-3]C*R[-3]C[2]"
            End With

            '###################################################################
            '#############   TOTAL HORAS EXTRAS,RECARGO NOCTURNO
            '###################################################################


            .Rows(CStr(filCelFn - 2) & ":" & CStr(filCelFn - 2)).RowHeight = 31.5

            With .Range(.Cells(filCelInc, 1).Offset(17, 1), .Cells(filCelInc, 1).Offset(17, 5))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 14
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "TOTAL POR TRABAJO ESTR. Y NOCTURNO:"
            End With

            With .Range(.Cells(filCelInc, 1).Offset(17, 7), .Cells(filCelInc, 1).Offset(17, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 14
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-2]C[-6]+R[-2]C[-1]"
            End With
        End With
    End Sub
    Public Sub ajusteHojaHrExt()
        With Globals.ThisAddIn.Application
            .Columns("A:A").ColumnWidth = 6
            .Columns("K:K").ColumnWidth = 6

            .Columns("C:C").ColumnWidth = 3
            .Columns("H:H").ColumnWidth = 3

            .Columns("F:F").ColumnWidth = 15
        End With
    End Sub
    Public Sub contenedorHrExt()
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

    Function resultadoExtNoc(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultadoExtNoc = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(17, 7).Address
        End With
    End Function

End Module
