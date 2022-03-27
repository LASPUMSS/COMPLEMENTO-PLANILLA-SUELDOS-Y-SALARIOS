Module M08_RESUMEN_PLN_TRIBUTARIA
    Public Sub plantillaResumenPlnTrib(ByVal gestionPlanilla As Long, ByVal mesPlanilla As Integer)
        With Globals.ThisAddIn.Application

            .Cells(1, 1).Value = "PLANILLA DE RETENCIONES DEL RC-IVA CORRESPONDIENTE AL MES DE " & mesTexto(mesPlanilla) & " DEL " & CStr(gestionPlanilla)
            .Cells(2, 1).Value = "(EXPRESADO EN BOLIVIANOS)"

            .Cells(4, 19).Value = "SALDO A FAVOR"
            .Cells(4, 21).Value = "SALDO A FAVOR DE DEPENDIENTE DEL PERIODO ANTERIOR"

            .Cells(5, 2).Value = "N°"
            .Cells(5, 3).Value = "AÑO "
            .Cells(5, 4).Value = "PERIODO"
            .Cells(5, 5).Value = "CÓDIGO DEPENDIENTE"
            .Cells(5, 6).Value = "NOMBRES"
            .Cells(5, 7).Value = "PRIMER APELLIDO"
            .Cells(5, 8).Value = "SEGUNDO APELLIDO"
            .Cells(5, 9).Value = "NÚMERO DOCUMENTO DE IDENTIDAD "
            .Cells(5, 10).Value = "TIPO DE DOCUMENTO"
            .Cells(5, 11).Value = "NOVEDADES I-INCORPORACIÓN V-VIGENTE D-DESVINCULADO"
            .Cells(5, 12).Value = "MONTO DE INGRESO NETO"
            .Cells(5, 13).Value = "DOS(2) S.M.N NO IMPONIBLE"
            .Cells(5, 14).Value = "IMPORTE SUJETO A IMPUESTO BASE IMPONIBLE"
            .Cells(5, 15).Value = "IMPUESTO RC-IVA"
            .Cells(5, 16).Value = "FORM 110"
            .Cells(5, 17).Value = "13% DE DOS (2) S.M.N."

            .Cells(5, 18).Value = "FISCO"
            .Cells(5, 19).Value = "DEPENDIENTE"

            .Cells(5, 20).Value = "MES ANTERIOR"
            .Cells(5, 21).Value = "MANTENIMIENTO DE VALOR"
            .Cells(5, 22).Value = "SUB-TOTAL"

            .Cells(5, 23).Value = "SALDO TOTAL A FAVOR DEL DEPENDIENTE"
            .Cells(5, 24).Value = "SALDO UTILIZADO"
            .Cells(5, 25).Value = "LIQUIDACIÓN DE LAS RETENCIONES"
            .Cells(5, 26).Value = "SALDO A FAVOR DEPENDIENTE P/MES SGTE."

            .Cells(6, 12).Value = "A"
            .Cells(6, 13).Value = "B"
            .Cells(6, 14).Value = "C"
            .Cells(6, 15).Value = "D"


            .Cells(7, 2).Value = "-"
            .Cells(7, 3).Value = "-"
            .Cells(7, 4).Value = "-"
            .Cells(7, 5).Value = "-"
            .Cells(7, 6).Value = "-"
            .Cells(7, 7).Value = "-"
            .Cells(7, 8).Value = "-"
            .Cells(7, 9).Value = "-"
            .Cells(7, 10).Value = "-"
            .Cells(7, 11).Value = "-"
            .Cells(7, 12).Value = "A=TG-AFP"
            .Cells(7, 13).Value = "B=2 S.M.N."
            .Cells(7, 14).Value = "C=A-B"
            .Cells(7, 15).Value = "D=C*13%"
            .Cells(7, 16).Value = "-"
            .Cells(7, 17).Value = "-"
            .Cells(7, 18).Value = "-"
            .Cells(7, 19).Value = "-"
            .Cells(7, 20).Value = "-"
            .Cells(7, 21).Value = "-"
            .Cells(7, 22).Value = "-"
            .Cells(7, 23).Value = "-"
            .Cells(7, 24).Value = "-"
            .Cells(7, 25).Value = "-"
            .Cells(7, 26).Value = "-"

            With .Range(.Cells(1, 1), .Cells(1, 27))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 48
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(2, 1), .Cells(2, 27))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 26
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(6, 2), .Cells(6, 27))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 9
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
            End With

            With .Range(.Cells(7, 2), .Cells(7, 26))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 9
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.599993896298105
            End With

            With .Rows("4:4")
                .RowHeight = 50
                .WrapText = True
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
                .Font.Size = 9
            End With

            With .Rows("5:5")
                .RowHeight = 127
                .WrapText = True
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Font.TintAndShade = -0.249977111117893
                .Font.Size = 9
            End With

            .Range(.Cells(4, 18), .Cells(4, 19)).Select()
            .Selection.Merge
            formatoTablas()

            .Range(.Cells(4, 20), .Cells(4, 22)).Select()
            .Selection.Merge
            formatoTablas()


            .Range(.Cells(5, 2), .Cells(7, 26)).Select()
            formatoTablas()

            .Rows("7:7").RowHeight = 40

            .Columns("A:A").ColumnWidth = 9
            .Columns("B:B").ColumnWidth = 7
            .Columns("C:C").ColumnWidth = 10
            .Columns("D:D").ColumnWidth = 15
            .Columns("E:E").ColumnWidth = 15
            .Columns("F:F").ColumnWidth = 15
            .Columns("G:G").ColumnWidth = 15
            .Columns("H:H").ColumnWidth = 15
            .Columns("I:I").ColumnWidth = 15
            .Columns("J:J").ColumnWidth = 15
            .Columns("K:K").ColumnWidth = 15
            .Columns("L:L").ColumnWidth = 15
            .Columns("M:M").ColumnWidth = 15
            .Columns("N:N").ColumnWidth = 15
            .Columns("O:O").ColumnWidth = 15
            .Columns("P:P").ColumnWidth = 15
            .Columns("Q:Q").ColumnWidth = 15
            .Columns("R:R").ColumnWidth = 15
            .Columns("S:S").ColumnWidth = 15
            .Columns("T:T").ColumnWidth = 15
            .Columns("U:U").ColumnWidth = 15
            .Columns("V:V").ColumnWidth = 15
            .Columns("W:W").ColumnWidth = 15
            .Columns("X:X").ColumnWidth = 15
            .Columns("Y:Y").ColumnWidth = 15
            .Columns("Z:Z").ColumnWidth = 15
            .Columns("AA:AA").ColumnWidth = 15
            .Columns("AB:AB").ColumnWidth = 9
        End With
    End Sub


    Public Sub CalculosPlnTrib(ByVal hojaPrincipalDatos As Excel.Worksheet, ByVal hojaPlnTrb As Excel.Worksheet, ByVal SALARIO_MINIMO_NACINAL_VIGENTE_02 As String)

        With Globals.ThisAddIn.Application
            Dim Celda As Excel.Range
            Dim CI_BUSCADO As Excel.Range
            Dim n1 As Long
            Dim n2 As Long

            '################################################################
            '#############    VARIABLES DATOS
            '################################################################

            Dim GESTION As String
            Dim PERIODO As String
            Dim CODIGO_DEPENDIENTE As String
            Dim NOMBRES As String
            Dim PRIMER_APELLIDO As String
            Dim SEGUNDO_APELLIDO As String

            Dim TIPO_DE_DOCUMENTO As String
            Dim NOVEDADES_INCORPORACION_VIGENTE_DESVINCULADO As String

            Dim DOS_SMN_NO_IMPONIBLE As String
            Dim IMPORTE_SUJETO_A_IMPUESTO_BASE_IMPONIBLE As String
            Dim IMPUESTO_RC_IVA As String
            Dim FORM_110 As String
            Dim TRECE_PORCIENTO_DE_DOS_SMN As String

            Dim FISCO As String
            Dim DEPENDIENTE As String
            Dim MES_ANTERIOR As String
            Dim MANTENIMIENTO_DE_VALOR As String
            Dim SUB_TOTAL As String

            Dim SALDO_TOTAL_A_FAVOR_DEL_DEPENDIENTE As String
            Dim SALDO_UTILIZADO As String
            Dim LIQUIDACION_DE_LAS_RETENCIONES As String
            Dim SALDO_A_FAVOR_DEPENDIENTE_P_MES_SGTE As String


            hojaPrincipalDatos.Activate()
            n2 = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            hojaPlnTrb.Activate()
            n1 = .Cells(.Rows.Count, 9).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row


            hojaPlnTrb.Activate()
            .Range(.Cells(8, 9), .Cells(n1, 9)).Select()

            For Each Celda In .Selection

                hojaPrincipalDatos.Activate()

                Dim CI_HOJA_PLN_TRB As String
                CI_HOJA_PLN_TRB = Celda.Value

                CI_BUSCADO = .Range(.Cells(16, 2), .Cells(n2, 2)).Find(CI_HOJA_PLN_TRB)

                GESTION = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(4, 0).Address)
                PERIODO = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(5, 0).Address)
                CODIGO_DEPENDIENTE = resultadoUb(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 27).Address)
                NOMBRES = resultadoUb(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 24).Address)
                PRIMER_APELLIDO = resultadoUb(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 25).Address)
                SEGUNDO_APELLIDO = resultadoUb(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 26).Address)

                TIPO_DE_DOCUMENTO = resultadoUb(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 28).Address)
                NOVEDADES_INCORPORACION_VIGENTE_DESVINCULADO = "=IF(" & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 29).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & " =""INCORPORADO"",""I"","""") " &
                                                            " & IF( " & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 29).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "=""VIGENTE"",""V"","""")" &
                                                            " & IF( " & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 29).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & " =""DESVINCULADO"",""D"","""") "

                'DOS_SMN_NO_IMPONIBLE = "=ROUND(2*" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0)"
                DOS_SMN_NO_IMPONIBLE = "=IF(RC[-1]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                                "),ROUND(2*" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0),0)"
                'IMPORTE_SUJETO_A_IMPUESTO_BASE_IMPONIBLE = "=ROUND(RC[-2]-RC[-1],0)"
                IMPORTE_SUJETO_A_IMPUESTO_BASE_IMPONIBLE = "=IF(RC[-2]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                                                    "),ROUND(RC[-2]-RC[-1],0),0)"
                'IMPUESTO_RC_IVA = "=ROUND(RC[-1]*13%,0)"
                IMPUESTO_RC_IVA = "=IF(RC[-3]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                           "),ROUND(RC[-1]*13%,0),0)"
                'FORM_110 = "=ROUND(" & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 30).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0)"
                FORM_110 = "=IF(RC[-4]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                    "),ROUND(" & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 30).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0),0)"
                'TRECE_PORCIENTO_DE_DOS_SMN = "=ROUND(13%*2*" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0)"
                TRECE_PORCIENTO_DE_DOS_SMN = "=IF(RC[-5]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                                     "),ROUND(13%*2*" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0),0)"

                FISCO = "=ROUND(IF(RC[-3]>(RC[-2]+RC[-1]),(RC[-3]-RC[-2]-RC[-1]),0),0)"
                DEPENDIENTE = "=ROUND(IF((RC[-3]+RC[-2])>RC[-4],(RC[-3]+RC[-2]-RC[-4]),0),0)"
                'MES_ANTERIOR = "=ROUND(" & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 31).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0)"
                MES_ANTERIOR = "=IF(RC[-8]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                        "),ROUND(" & resultadoUb2(hojaPrincipalDatos.Name, CI_BUSCADO.Offset(0, 31).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ",0),0)"
                'MANTENIMIENTO_DE_VALOR = "=ROUND(((" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(8, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "/" &
                'resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(9, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ")-1)*RC[-1],0)"
                MANTENIMIENTO_DE_VALOR = "=IF(RC[-9]>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                                    "),ROUND(((" & resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(8, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "/" &
                                     resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(9, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & ")-1)*RC[-1],0),0)"
                SUB_TOTAL = "=ROUND(RC[-2]+RC[-1],0)"

                SALDO_TOTAL_A_FAVOR_DEL_DEPENDIENTE = "=ROUND(RC[-1]+RC[-4],0)"
                SALDO_UTILIZADO = "=IF(RC[-6]>RC[-1],RC[-1],0)+IF(AND(RC[-6]>0,RC[-1]>RC[-6]),RC[-6],0)"
                LIQUIDACION_DE_LAS_RETENCIONES = "=ROUND(RC[-7]-RC[-1],0)"
                SALDO_A_FAVOR_DEPENDIENTE_P_MES_SGTE = "=ROUND(RC[-3]-RC[-2],0)"

                hojaPlnTrb.Activate()

                .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Row - 7
                .Cells(.Rows.Count, 3).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = GESTION
                .Cells(.Rows.Count, 4).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = PERIODO
                .Cells(.Rows.Count, 5).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = CODIGO_DEPENDIENTE
                .Cells(.Rows.Count, 6).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = NOMBRES
                .Cells(.Rows.Count, 7).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = PRIMER_APELLIDO
                .Cells(.Rows.Count, 8).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = SEGUNDO_APELLIDO

                .Cells(.Rows.Count, 10).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = TIPO_DE_DOCUMENTO
                .Cells(.Rows.Count, 11).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = NOVEDADES_INCORPORACION_VIGENTE_DESVINCULADO

                .Cells(.Rows.Count, 13).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = DOS_SMN_NO_IMPONIBLE
                .Cells(.Rows.Count, 14).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = IMPORTE_SUJETO_A_IMPUESTO_BASE_IMPONIBLE
                .Cells(.Rows.Count, 15).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = IMPUESTO_RC_IVA
                .Cells(.Rows.Count, 16).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = FORM_110
                .Cells(.Rows.Count, 17).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = TRECE_PORCIENTO_DE_DOS_SMN

                .Cells(.Rows.Count, 18).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = FISCO
                .Cells(.Rows.Count, 19).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = DEPENDIENTE
                .Cells(.Rows.Count, 20).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = MES_ANTERIOR
                .Cells(.Rows.Count, 21).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = MANTENIMIENTO_DE_VALOR
                .Cells(.Rows.Count, 22).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = SUB_TOTAL

                .Cells(.Rows.Count, 23).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = SALDO_TOTAL_A_FAVOR_DEL_DEPENDIENTE
                .Cells(.Rows.Count, 24).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = SALDO_UTILIZADO
                .Cells(.Rows.Count, 25).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = LIQUIDACION_DE_LAS_RETENCIONES
                .Cells(.Rows.Count, 26).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = SALDO_A_FAVOR_DEPENDIENTE_P_MES_SGTE

            Next

            .Columns("L:Z").NumberFormat = "#,##0"
            .Columns("B:K").HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            .Range(.Cells(8, 2), .Cells(n1, 26)).Select()
            formatoTablas()
        End With

    End Sub
    Public Sub recogerResRC_IVA(ByVal hojaPrePlanilla As Excel.Worksheet, ByVal hojaPlnTrb As Excel.Worksheet)
        With Globals.ThisAddIn.Application
            Dim Celda As Excel.Range
            Dim n1 As Long
            Dim n2 As Long
            Dim i As Long

            hojaPrePlanilla.Activate()
            n2 = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            hojaPlnTrb.Activate()
            n1 = .Cells(.Rows.Count, 9).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row



            hojaPlnTrb.Activate()
            .Range(.Cells(8, 9), .Cells(n1, 9)).Select

            For Each Celda In .Selection

                Dim RC_IVA As String
                Dim CI_HOJA_PLN_TRB As String

                RC_IVA = resultadoUb(hojaPlnTrb.Name, Celda.Offset(0, 16).Address)

                CI_HOJA_PLN_TRB = Celda.Value

                hojaPrePlanilla.Activate()

                For i = 2 To n2
                    If .Cells(i, 2).Value = CI_HOJA_PLN_TRB Then

                        .Cells(i, 2).Offset(0, 18).Value = RC_IVA
                    End If
                Next

            Next
        End With
    End Sub

End Module
