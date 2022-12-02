Module M05_DetalleDominicalDomigTrab
    Public domingosMes As String

    Public Sub detalleDominicalDomTrab(ByVal CI As String,
                            ByVal nombreCompleto As String,
                            ByVal haberBasico As String,
                            ByVal DOMINICAL As String,
                            ByVal funcion As String,
                            ByVal condicionDom As String,
                            ByVal hrsDomTrabajados As String,
                            ByVal gestionPlanilla As String,
                            ByVal mesPlanilla As String,
                            ByVal diasPlanilla As String)

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim filCelInc As Long
            Dim filCelFn As Long
            Dim rangoHrExt As Excel.Range

            ajusteHojaDominicalDomTrab()

            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(2, 0).Row
            filCelInc = n + 1
            filCelFn = n + 20

            rangoHrExt = .Range(.Cells(n, 1), .Cells(filCelFn, 11))
            rangoHrExt.Select()

            contenedorDominicalDomTrab()

            '##############################################################
            '#############    TITULO DE REPORTE HORAS EXTRAORDINARIAS
            '##############################################################

            .Cells(n, 1).Value = "DOMINICAL Y DOMINGOS TRABAJADOS"
            .Cells(filCelFn, 1).Value = "CONFORME AL ARTICULO 23 DEL D.S. 36091."

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
                .Value = "C.I. :"
                .Offset(0, 2).Value = CI
                .Offset(0, 2).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 1), .Cells(filCelInc, 1).Offset(2, 1))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "DOMINICAL:"
                .Offset(0, 2).Value = DOMINICAL
            End With
            With .Range(.Cells(filCelInc, 1).Offset(0, 6), .Cells(filCelInc, 1).Offset(0, 6))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "HABER BASICO:"
                .Offset(0, 2).Value = haberBasico
                .Offset(0, 2).NumberFormat = "#,##0.00"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(1, 6), .Cells(filCelInc, 1).Offset(1, 6))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "FUNCION:"
                .Offset(0, 2).Value = funcion
            End With
            With .Range(.Cells(filCelInc, 1).Offset(2, 6), .Cells(filCelInc, 1).Offset(2, 6))
                .Merge
                .Font.Bold = True
                .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .Value = "COND. DOM. :"
                .Offset(0, 2).Value = condicionDom
            End With
            '###################################################################
            '#############   TITULOS DOMINGOS TRABAJADOS Y DOMINICAL
            '###################################################################

            With .Range(.Cells(filCelInc, 1).Offset(4, 1), .Cells(filCelInc, 1).Offset(4, 4))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "CALCULO DOMINGO TRABAJADO"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(4, 6), .Cells(filCelInc, 1).Offset(4, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "CALCULO DOMINICAL"
            End With
            '###################################################################
            '#############   ETIQUETAS DOMINGOS TRABAJADOS
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

            'FILA TRIPLE

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
                .Value = "3"
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
                .Value = hrsDomTrabajados
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
                .Value = "TOTAL POR DOM. TRAB."
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
            '#############   ETIQUETAS DOMINICAL
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
                .Value = "=" & diasPlanilla
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 7), .Cells(filCelInc, 1).Offset(6, 7))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "-"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(6, 8), .Cells(filCelInc, 1).Offset(6, 8))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .FormulaR1C1 = $"=INT( (({hojaPrincipalDatos.Name}!R10C8 & ""/"" & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8)-((1) & ""/""  & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8) + WEEKDAY(((1) & ""/""  & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8)-1 ))/7)"
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

            'FILA DIAS

            With .Range(.Cells(filCelInc, 1).Offset(8, 6), .Cells(filCelInc, 1).Offset(8, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=IF(AND(R[-6]C[-3]=""CUMPLE"",R[-6]C[2]=""LLEGO PUNTUAL"",R[-7]C[2]=""OBRERO""),R[-8]C[2],0)"
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
                .Value = "=IF(R[-2]C[-2]-R[-2]C<=0,1,R[-2]C[-2]-R[-2]C)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(8, 9), .Cells(filCelInc, 1).Offset(8, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "DIAS"
            End With

            'FILA DOMINGOS MES
            With .Range(.Cells(filCelInc, 1).Offset(12, 6), .Cells(filCelInc, 1).Offset(12, 6))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "#,##0.00"
                .FormulaR1C1 = "=R[-4]C/R[-4]C[2]"
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
                .NumberFormat = "#,##0"
                .FormulaR1C1 = $"=INT( (({hojaPrincipalDatos.Name}!R10C8 & ""/"" & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8)-((1) & ""/""  & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8) + WEEKDAY(((1) & ""/""  & {hojaPrincipalDatos.Name}!R9C8 & ""/""  & {hojaPrincipalDatos.Name}!R8C8)-1 ))/7)"
            End With
            With .Range(.Cells(filCelInc, 1).Offset(12, 9), .Cells(filCelInc, 1).Offset(12, 9))
                .Merge
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorLight2
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Value = "DOM. MES."
            End With

            ' TOTAL DOMINICAL

            With .Range(.Cells(filCelInc, 1).Offset(14, 6), .Cells(filCelInc, 1).Offset(14, 9))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
                .Value = "TOTAL DOMINICAL"
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
            '#############   TOTAL DOMINICAL Y HORAS DOMINGO
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
                .Value = "TOTAL POR DOM. TRAB. Y DOMINICAL:"
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
    Public Sub ajusteHojaDominicalDomTrab()
        With Globals.ThisAddIn.Application
            .Columns("A:A").ColumnWidth = 6
            .Columns("K:K").ColumnWidth = 6

            .Columns("C:C").ColumnWidth = 3
            .Columns("H:H").ColumnWidth = 3

            .Columns("F:F").ColumnWidth = 15
        End With
    End Sub
    Public Sub contenedorDominicalDomTrab()
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

    Public Function resultadoDomTrbDom(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultadoDomTrbDom = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(17, 7).Address
        End With
    End Function

End Module
