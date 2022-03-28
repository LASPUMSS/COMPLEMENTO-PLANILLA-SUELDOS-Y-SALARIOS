Module M09_RESUMEN_APORTE_PATRONAL
    Public Sub plantillaResumenAportePatronal(ByVal gestionPlanilla As Long, ByVal mesPlanilla As Integer)
        With Globals.ThisAddIn.Application

            .DisplayAlerts = False

            .Cells(1, 1).Value = "PLANILLA DE APORTES PATRONALES Y BENEFICIOS SOCIALES"
            .Cells(2, 1).Value = "CORRESPONDIENTE AL MES DE " & mesTexto(mesPlanilla) & " DE " & CStr(gestionPlanilla)
            .Cells(3, 1).Value = "(EXPRESADO EN BOLIVIANOS)"

            With .Range(.Cells(1, 1), .Cells(1, 14))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 12
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With
            With .Range(.Cells(2, 1), .Cells(2, 14))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 12
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With
            With .Range(.Cells(3, 1), .Cells(3, 14))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 12
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With

            .Cells(5, 6).Value = "APORTES PATRONALES"
            .Cells(5, 7).Value = "BENEFICIOS SOCIALES"

            With .Range(.Cells(5, 6), .Cells(5, 10))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 12
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = 0
            End With
            With .Range(.Cells(5, 11), .Cells(5, 13))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 12
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = 0
            End With

            .Range(.Cells(5, 6), .Cells(5, 13)).Select()
            formatoTablas()

            .Cells(6, 2) = "N°"
            .Cells(6, 3) = "C.I."
            .Cells(6, 4) = "NOMBRE Y APELLIDO"
            .Cells(6, 5) = "TOTAL GANADO"
            .Cells(6, 6) = "C.N.S. 10%"
            .Cells(6, 7) = "A.F.P. 1,71%"
            .Cells(6, 8) = "PROVIVIENDA 2%"
            .Cells(6, 9) = "APORTE SOL. 3%"
            .Cells(6, 10) = "TOTAL 16,71%"
            .Cells(6, 11) = "AGUINALDO 8,33333%"
            .Cells(6, 12) = "INDEM. 8,33333%"
            .Cells(6, 13) = "TOTAL"


            With .Range(.Cells(6, 2), .Cells(6, 13))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 10
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = 0
                .WrapText = True
            End With

            .Rows("6:6").RowHeight = 40
            .Range(.Cells(6, 2), .Cells(6, 13)).Select()
            formatoTablas()

        End With
    End Sub

    Public Sub CalculosAportePatronal(ByVal hojaPrePlanilla As Excel.Worksheet, ByVal hojaAportePatronal As Excel.Worksheet)
        With Globals.ThisAddIn.Application

            .DisplayAlerts = False

            Dim n As Long
            Dim Celda As Excel.Range

            hojaPrePlanilla.Activate()
            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(2, 2), .Cells(n, 2)).Select()

            For Each Celda In .Selection

                hojaAportePatronal.Activate()
                .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, -1).Address)
                .Cells(.Rows.Count, 3).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Address)
                .Cells(.Rows.Count, 4).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, 1).Address)
                .Cells(.Rows.Count, 5).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, 16).Address)

                .Cells(.Rows.Count, 6).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-1]*10%"
                .Cells(.Rows.Count, 7).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]*1.71%"
                .Cells(.Rows.Count, 8).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-3]*2%"
                .Cells(.Rows.Count, 9).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-4]*3%"
                .Cells(.Rows.Count, 10).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"
                .Cells(.Rows.Count, 11).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=(1/12)*RC[-6]"
                .Cells(.Rows.Count, 12).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=(1/12)*RC[-7]"
                .Cells(.Rows.Count, 13).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]+RC[-1]"

            Next

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(7, 5), .Cells(n, 13)).NumberFormat = "#,##0.00"
            .Range(.Cells(7, 2), .Cells(n, 13)).Select()
            formatoTablas()
            .Cells(7, 2).Select()
        End With
    End Sub

    Public Sub CopiarResulatadosAportePatronal(ByVal hojaPrePlanilla As Excel.Worksheet, ByVal hojaAportePatronal As Excel.Worksheet)
        With Globals.ThisAddIn.Application

            Dim n As Long
            Dim Celda As Excel.Range

            hojaAportePatronal.Activate()

            n = .Cells(.Rows.Count, 10).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(7, 10), .Cells(n, 10)).Select()

            For Each Celda In .Selection

                hojaPrePlanilla.Activate()
                .Cells(.Rows.Count, 27).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaAportePatronal.Name, Celda.Address)
                .Cells(.Rows.Count, 28).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaAportePatronal.Name, Celda.Offset(0, 3).Address)
                .Cells(.Rows.Count, 29).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]+RC[-1]+RC[-11]"
                .Cells(.Rows.Count, 30).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-11]+RC[-10]+RC[-3]"
                .Cells(.Rows.Count, 31).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]-RC[-1]"
                .Cells(.Rows.Count, 32).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]/RC[-3]"
                .Cells(.Rows.Count, 33).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[-2]/RC[-4]"

            Next

            .Cells(1, 27).CurrentRegion.Select()
            .Selection.NumberFormat = "#,##0.00"
            formatoTablas()


        End With
    End Sub

End Module
