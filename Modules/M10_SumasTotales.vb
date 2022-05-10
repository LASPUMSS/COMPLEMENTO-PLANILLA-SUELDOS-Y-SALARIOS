Module M10_SumasTotales
    Public Sub SumasToalesGeneral(ByVal hojaPrePlanilla As Excel.Worksheet, ByVal hojaResumenAFP As Excel.Worksheet, ByVal hojaResumenRC_IVA As Excel.Worksheet, ByVal hojaAportePatronal As Excel.Worksheet)

        hojaPrePlanilla.Activate()
        SumasTotalesPrePlanilla()

        hojaResumenAFP.Activate()
        SumasTotalesAFP()

        hojaResumenRC_IVA.Activate()
        SumasTotalesPlanillaTrib()

        hojaAportePatronal.Activate()
        SumasTotalesAportePatronal()

        hojaPrePlanilla.Activate()

    End Sub

    Public Sub SumasTotalesPrePlanilla()

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Integer
            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            With .Range(.Cells(n + 1, 1), .Cells(n + 1, 10))
                .Merge
                .Value = "TOTAL"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Rows(CStr(n + 1) & ":" & CStr(n + 1))
                .RowHeight = 20
                .Font.Bold = True
                .Font.Size = 13
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            .Range(.Cells(n + 1, 1), .Cells(n + 1, 24)).Select()
            formatoTablas()

            .Range(.Cells(n + 1, 27), .Cells(n + 1, 33)).Select()
            formatoTablas()

            For i = 11 To 23
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 1 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            For i = 27 To 31
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 1 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            For i = 32 To 33
                .Cells(n + 1, i).FormulaR1C1 = "=AVERAGE(R[-" & n - 1 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            .Cells.EntireColumn.AutoFit()
        End With
    End Sub


    Public Sub SumasTotalesAFP()
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Integer
            n = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row


            With .Range(.Cells(n + 1, 1), .Cells(n + 1, 5))
                .Merge
                .Value = "TOTAL"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Rows(CStr(n + 1) & ":" & CStr(n + 1))
                .RowHeight = 20
                .Font.Bold = True
                .Font.Size = 13
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Range(.Cells(n + 1, 1), .Cells(n + 1, 15))
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(n + 1, 17), .Cells(n + 1, 20))
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent6
                .Interior.TintAndShade = -0.249977111117893
            End With

            .Range(.Cells(n + 1, 1), .Cells(n + 1, 15)).Select()
            formatoTablas()

            .Range(.Cells(n + 1, 17), .Cells(n + 1, 20)).Select()
            formatoTablas()

            For i = 6 To 15
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 3 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            For i = 17 To 20
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 3 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            .Cells.EntireColumn.AutoFit()
        End With

    End Sub

    Public Sub SumasTotalesAportePatronal()

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Integer
            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            With .Range(.Cells(n + 1, 2), .Cells(n + 1, 4))
                .Merge
                .Value = "TOTAL"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Rows(CStr(n + 1) & ":" & CStr(n + 1))
                .RowHeight = 20
                .Font.Bold = True
                .Font.Size = 13
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Range(.Cells(n + 1, 2), .Cells(n + 1, 13))
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = 0
            End With
            .Range(.Cells(n + 1, 2), .Cells(n + 1, 13)).Select()
            formatoTablas()

            For i = 5 To 13
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 6 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            .Cells.EntireColumn.AutoFit()
        End With

    End Sub

    Public Sub SumasTotalesPlanillaTrib()

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Integer
            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            With .Range(.Cells(n + 1, 2), .Cells(n + 1, 11))
                .Merge
                .Value = "TOTAL"
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Rows(CStr(n + 1) & ":" & CStr(n + 1))
                .RowHeight = 20
                .Font.Bold = True
                .Font.Size = 13
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            End With

            With .Range(.Cells(n + 1, 2), .Cells(n + 1, 26))
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
            End With

            .Range(.Cells(n + 1, 2), .Cells(n + 1, 26)).Select()
            formatoTablas()

            For i = 12 To 26
                .Cells(n + 1, i).FormulaR1C1 = "=SUM(R[-" & n - 7 & "]C:R[-1]C)"
                .Cells(n + 1, i).NumberFormat = "#,##0.00"
            Next

            .Cells.EntireColumn.AutoFit()
        End With
    End Sub

End Module
