Module M07_RESUMEN_AFP
    Public Sub hojaANS()

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim Celda As Excel.Range
            Dim hojaInformeDetallado As Excel.Worksheet
            Dim hojaResumenANS As Excel.Worksheet

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            hojaResumenANS = .ActiveSheet
            'Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

            .Sheets.Add()
            hojaInformeDetallado = .ActiveSheet

            hojaResumenANS.Activate()
            .Range(.Cells(4, 2), .Cells(n, 2)).Select

            For Each Celda In .Selection

                Dim totalGanado As String
                Dim nombreCompleto As String
                Dim carnet As String

                totalGanado = "=" & hojaResumenANS.Name & "!" & Celda.Offset(0, 2).Address
                nombreCompleto = "=" & hojaResumenANS.Name & "!" & Celda.Offset(0, 1).Address
                carnet = "=" & hojaResumenANS.Name & "!" & Celda.Address

                hojaInformeDetallado.Activate()

                APN(totalGanado, nombreCompleto, carnet)


                Celda.Offset(0, 9).Value = resultado01(hojaInformeDetallado.Name)
                Celda.Offset(0, 10).Value = resultado02(hojaInformeDetallado.Name)
                Celda.Offset(0, 11).Value = resultado03(hojaInformeDetallado.Name)
                Celda.Offset(0, 12).FormulaR1C1 = "=SUM(RC[-3]:RC[-1])"
                Celda.Hyperlinks.Add(Anchor:=Celda, Address:="", SubAddress:=vinculo(hojaInformeDetallado.Name))

                ' AFP
                Celda.Offset(0, 4).FormulaR1C1 = "=RC[-2]*10%"
                Celda.Offset(0, 5).FormulaR1C1 = "=RC[-3]*1.71%"
                Celda.Offset(0, 6).FormulaR1C1 = "=RC[-4]*0.5%"
                Celda.Offset(0, 7).FormulaR1C1 = "=RC[-5]*0.5%"
                Celda.Offset(0, 8).FormulaR1C1 = "=SUM(RC[-4]:RC[-1])"

                'TOTAL CONTRIBUCIÓN DEL EMPLEADO
                Celda.Offset(0, 13).FormulaR1C1 = "=RC[-5]+RC[-1]"


                '###############################################
                '##########  CONTRIBUCIONES DEL EMPLEADOR
                '###############################################
                Celda.Offset(0, 15).FormulaR1C1 = "=RC[-13]*1.71%"
                Celda.Offset(0, 16).FormulaR1C1 = "=RC[-14]*3%"
                Celda.Offset(0, 17).FormulaR1C1 = "=RC[-15]*2%"
                Celda.Offset(0, 18).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"

            Next

            hojaResumenANS.Activate()

            'DAR FORMATO A LA TABLA DE CONTRIBUCIONES DE EMPLEADOR
            .Cells(1, 17).CurrentRegion.Select
            formatoTablas()

            .Cells(4, 17).Select
            .Range(.Selection, .Selection.End(Microsoft.Office.Interop.Excel.XlDirection.xlToRight)).Select
            .Range(.Selection, .Selection.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown)).Select

            With .Selection
                .NumberFormat = "#,##0.00"
            End With
        End With

    End Sub
    Public Sub plantillaResumenAFP()
        With Globals.ThisAddIn.Application
            .Cells(1, 1).Value = "CALCULO DE LAS APORTACIONES DEL TRABAJADOR"
            .Cells(1, 17).Value = "CALCULO DE CONT. DEL EMPLEADOR"

            .Rows("2:2").RowHeight = 43.5

            .Cells(2, 1).Value = "N°"
            .Cells(2, 2).Value = "C.I."
            .Cells(2, 3).Value = "APELLIDOS Y NOMBRES"
            .Cells(2, 4).Value = "TOTAL GANADO"
            .Cells(2, 5).Value = "EDAD"
            .Cells(2, 6).Value = "COTIZACIÓN MENSUAL"
            .Cells(2, 7).Value = "PRIMA RIESGO COMUN"
            .Cells(2, 8).Value = "COMISIÓN APF"
            .Cells(2, 9).Value = "AP. LABORAL SOLIDARIO"
            .Cells(2, 10).Value = "TOTAL AP. AFP"
            .Cells(2, 11).Value = "ANS PARA TG>13000"
            .Cells(2, 12).Value = "ANS PARA TG>25000"
            .Cells(2, 13).Value = "ANS PARA TG>35000"
            .Cells(2, 14).Value = "TOTAL AP. NAL. SOLIDARIO"
            .Cells(2, 15).Value = "TOTAL CONTRIB. MENSUAL"

            .Cells(2, 17).Value = "PRIMA RIESGO PROFESIONAL"
            .Cells(2, 18).Value = "APORTE PATRONAL SOLIDARIO"
            .Cells(2, 19).Value = "APORTE PATRONAL VIVIENDA"
            .Cells(2, 20).Value = "TOTAL APORTE PATRONAL"

            .Cells(3, 1).Value = "-"
            .Cells(3, 2).Value = "-"
            .Cells(3, 3).Value = "-"
            .Cells(3, 4).Value = "A"
            .Cells(3, 5).Value = "B"
            .Cells(3, 6).Value = "C=A*10%"
            .Cells(3, 7).Value = "D=A*1,71%"
            .Cells(3, 8).Value = "E=A*0,5%"
            .Cells(3, 9).Value = "F=A*0,5%"
            .Cells(3, 10).Value = "G=C+D+E+F"
            .Cells(3, 11).Value = "H"
            .Cells(3, 12).Value = "I"
            .Cells(3, 13).Value = "J"
            .Cells(3, 14).Value = "K=H+I+J"
            .Cells(3, 15).Value = "L=G+K"

            .Cells(3, 17).Value = "M=A*1,71%"
            .Cells(3, 18).Value = "N=A*3%"
            .Cells(3, 19).Value = "O=A*2%"
            .Cells(3, 20).Value = "P=M+N+O"

            With .Range(.Cells(1, 1), .Cells(1, 15))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 26
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Font.TintAndShade = 0
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(1, 17), .Cells(1, 20))
                .Merge
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 14
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent6
                .Interior.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(2, 1), .Cells(3, 15))
                .WrapText = True
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 9
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent1
                .Interior.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(2, 17), .Cells(3, 20))
                .WrapText = True
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 9
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent6
                .Interior.TintAndShade = -0.249977111117893
            End With
        End With
    End Sub
    Public Function resultado01(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultado01 = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(20, 7).Address
        End With
    End Function
    Public Function resultado02(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultado02 = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(20, 9).Address
        End With
    End Function
    Public Function resultado03(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            resultado03 = "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(20, 11).Address
        End With
    End Function
    Public Function vinculo(ByVal nombreHojaInformeDetallado As String) As String
        With Globals.ThisAddIn.Application
            Return "=" & nombreHojaInformeDetallado & "!" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Address & ":" & .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Offset(0, 12).Address
        End With
    End Function

End Module
