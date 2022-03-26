Module M02_HojaPrincipalParaPlanilla
    Public hojaPrincipalDatos As Excel.Worksheet
    Public hojaPrePlanilla As Excel.Worksheet
    Public hojaBonoAntDetallado As Excel.Worksheet
    Public hojaExtNoctDetallado As Excel.Worksheet
    Public hojaDomTrabDomDetallado As Excel.Worksheet
    Public hojaResumenAFP As Excel.Worksheet
    Public hojaResumenRC_IVA As Excel.Worksheet
    Public hojaAportePatronal As Excel.Worksheet

    Public Sub hojaPrinciplaPlanillaSueldosSalarios()
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim Celda As Excel.Range

            '################################################################
            '#############    VARIABLES DATOS ESPECIFICOS DE PRE PLANILLA
            '################################################################
            Dim TOTAL_HORAS_EXTR_NOCT As String
            Dim TOTAL_DOMINICAL_DOMIG_TRAB As String
            Dim TOTAL_BONO_ANTIGUEDAD As String

            '#############################################################
            '#############    VARIABLES DATOS GENERALES DE HOJA PRINCIPAL
            '#############################################################
            Dim NOMBRE_O_RAZON_SOCIAL As String
            Dim NIT As String
            Dim NUMERO_IDENTIFICADOR_MINISTERIO_TRABAJO As String
            Dim NUMERO_DE_EMPLEADOR_CAJA_DE_SALUD As String
            Dim GESTION_PLANILLA As String
            Dim MES_PLANILLA As String
            Dim DIA_PLANILLA As String
            Dim SALARIO_MINIMO_NACINAL_VIGENTE As String
            Dim SALARIO_MINIMO_NACINAL_VIGENTE_02 As String


            '################################################################
            '#############    VARIABLES DATOS ESPECIFICOS DE HOJA PRINCIPAL
            '################################################################
            Dim N_TABLA As Long
            Dim DOCUMENTO_DE_IDENTIDAD As String
            Dim APELLIDO_Y_NOMBRES As String
            Dim PAIS_DE_NACIONALIDAD As String
            Dim FECHA_DE_NACIMIENTO As String
            Dim FECHA_DE_INGRESO As String
            Dim SEXO_V_M As String
            Dim OCUPACION_QUE_DESEMP As String
            Dim HORAS_PAGADAS_DIA As String
            Dim DIAS_PAGADAS_MES As String
            Dim HABER_BASICO As String
            Dim BONO_PRODUCCION As String
            Dim SUBSIDIO_FRONTERAS As String
            Dim CONDICION_SUBSIDIO_FRONTERAS As String
            Dim HORAS_EXTRAORDINARIAS_TRABAJADAS As String
            Dim HORAS_NOCTURNAS_TRABAJADAS As String

            Dim TIPO_DE_TRABAJO_NOCTURNO As String
            Dim HORAS_TRABAJADOS_EN_DOMINGO As String
            Dim DOMINICAL As String
            Dim CONDICION_01_DOMINICAL As String
            Dim CONDICION_02_DOMINICAL As String
            Dim OTROS_BONOS As String
            Dim CONCEPTO_DE_OTRO_BONOS As String
            Dim OTROS_DESCUENTOS As String
            Dim CONCEPTO_DE_OTRO_DESCUENTOS As String


            hojaPrincipalDatos = .ActiveSheet
            'Sheets.Add After:=ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)

            .Sheets.Add()
            hojaPrePlanilla = .ActiveSheet

            .Sheets.Add()
            hojaBonoAntDetallado = .ActiveSheet

            .Sheets.Add()
            hojaExtNoctDetallado = .ActiveSheet

            .Sheets.Add()
            hojaDomTrabDomDetallado = .ActiveSheet

            .Sheets.Add()
            hojaResumenAFP = .ActiveSheet

            .Sheets.Add()
            hojaResumenRC_IVA = .ActiveSheet

            .Sheets.Add()
            hojaAportePatronal = .ActiveSheet

            hojaPrincipalDatos.Activate()

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(17, 2), .Cells(n, 2)).Select

            For Each Celda In .Selection

                hojaPrincipalDatos.Activate()

                NOMBRE_O_RAZON_SOCIAL = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Address)
                NIT = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(1, 0).Address)
                NUMERO_IDENTIFICADOR_MINISTERIO_TRABAJO = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(2, 0).Address)
                NUMERO_DE_EMPLEADOR_CAJA_DE_SALUD = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(3, 0).Address)
                GESTION_PLANILLA = resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(4, 0).Address)
                MES_PLANILLA = resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(5, 0).Address)
                DIA_PLANILLA = resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(6, 0).Address)
                SALARIO_MINIMO_NACINAL_VIGENTE = resultadoUb(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address)

                N_TABLA = Celda.Offset(0, -1).Value
                DOCUMENTO_DE_IDENTIDAD = resultadoUb(hojaPrincipalDatos.Name, Celda.Address)
                APELLIDO_Y_NOMBRES = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 1).Address)
                PAIS_DE_NACIONALIDAD = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 2).Address)
                FECHA_DE_NACIMIENTO = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 3).Address)
                FECHA_DE_INGRESO = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 4).Address)
                SEXO_V_M = "=IF(" & resultadoUb2(hojaPrincipalDatos.Name, Celda.Offset(0, 5).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "=""VARON"",""V"","""") & " &
                    "IF(" & resultadoUb2(hojaPrincipalDatos.Name, Celda.Offset(0, 5).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "=""MUJER"",""M"","""")"
                OCUPACION_QUE_DESEMP = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 6).Address)
                HORAS_PAGADAS_DIA = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 7).Address)
                DIAS_PAGADAS_MES = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 8).Address)
                HABER_BASICO = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 9).Address)
                BONO_PRODUCCION = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 10).Address)
                SUBSIDIO_FRONTERAS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 11).Address)
                CONDICION_SUBSIDIO_FRONTERAS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 12).Address)
                HORAS_EXTRAORDINARIAS_TRABAJADAS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 13).Address)
                HORAS_NOCTURNAS_TRABAJADAS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 14).Address)
                TIPO_DE_TRABAJO_NOCTURNO = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 15).Address)
                HORAS_TRABAJADOS_EN_DOMINGO = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 16).Address)

                DOMINICAL = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 17).Address)
                CONDICION_01_DOMINICAL = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 18).Address)
                CONDICION_02_DOMINICAL = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 19).Address)
                OTROS_BONOS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 20).Address)
                CONCEPTO_DE_OTRO_BONOS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 21).Address)
                OTROS_DESCUENTOS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 22).Address)
                CONCEPTO_DE_OTRO_DESCUENTOS = resultadoUb(hojaPrincipalDatos.Name, Celda.Offset(0, 23).Address)

                hojaPrePlanilla.Activate()
                generaHojaPlanillaPrePrincipal()

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = N_TABLA
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 1).Value = DOCUMENTO_DE_IDENTIDAD
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 2).Value = APELLIDO_Y_NOMBRES
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 3).Value = PAIS_DE_NACIONALIDAD
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 4).Value = FECHA_DE_NACIMIENTO
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 5).FormulaR1C1 = SEXO_V_M
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 6).Value = OCUPACION_QUE_DESEMP
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 7).Value = FECHA_DE_INGRESO
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 8).Value = HORAS_PAGADAS_DIA
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 9).Value = DIAS_PAGADAS_MES
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 10).Value = HABER_BASICO


                '########################################################################
                '#############    SECCIÓN BONO DE ANTIGUEDAD
                '########################################################################

                hojaBonoAntDetallado.Activate()

                Dim FECHA_PLANILLA As String
                FECHA_PLANILLA = "=" & DIA_PLANILLA & " & ""/"" & " & MES_PLANILLA & " & ""/"" & " & GESTION_PLANILLA

                BonoAntiguedad(DOCUMENTO_DE_IDENTIDAD, APELLIDO_Y_NOMBRES, SALARIO_MINIMO_NACINAL_VIGENTE, FECHA_DE_INGRESO, FECHA_PLANILLA)

                TOTAL_BONO_ANTIGUEDAD = resultadoBoAnt(hojaBonoAntDetallado.Name)

                hojaPrePlanilla.Activate()
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 11).Value = TOTAL_BONO_ANTIGUEDAD

                '########################################################################
                '#############    SECCIÓN BONO DE PRODUCCIÓN
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 12).Value = BONO_PRODUCCION


                '########################################################################
                '#############    SECCIÓN SECCIÓN DE FRONTERAS
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 13).Value = SUBSIDIO_FRONTERAS


                '########################################################################
                '#############    SECCIÓN HORAS EXTRAS Y RECARGO POR TRABAJO NOCTURNO
                '########################################################################

                hojaExtNoctDetallado.Activate()

                detallaHorasExtraordinarias(TIPO_DE_TRABAJO_NOCTURNO,
                                DOCUMENTO_DE_IDENTIDAD,
                                APELLIDO_Y_NOMBRES,
                                HABER_BASICO,
                                SEXO_V_M,
                                HORAS_EXTRAORDINARIAS_TRABAJADAS,
                                HORAS_NOCTURNAS_TRABAJADAS)

                TOTAL_HORAS_EXTR_NOCT = resultadoExtNoc(hojaExtNoctDetallado.Name)
                hojaPrePlanilla.Activate()
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 14).Value = TOTAL_HORAS_EXTR_NOCT

                '########################################################################
                '#############    SECCIÓN PAGO DOMICAL Y HORAS DOMINGOS TRABAJADOS
                '########################################################################


                hojaDomTrabDomDetallado.Activate()
                detalleDominicalDomTrab(DOCUMENTO_DE_IDENTIDAD,
                                APELLIDO_Y_NOMBRES,
                                HABER_BASICO,
                                DOMINICAL,
                                CONDICION_01_DOMINICAL,
                                CONDICION_02_DOMINICAL,
                                HORAS_TRABAJADOS_EN_DOMINGO,
                                GESTION_PLANILLA,
                                MES_PLANILLA,
                                DIA_PLANILLA)

                TOTAL_DOMINICAL_DOMIG_TRAB = resultadoDomTrbDom(hojaDomTrabDomDetallado.Name)
                hojaPrePlanilla.Activate()
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 15).Value = TOTAL_DOMINICAL_DOMIG_TRAB

                '########################################################################
                '#############    SECCIÓN OTROS BONOS
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 16).Value = OTROS_BONOS


                '########################################################################
                '#############    SECCIÓN TOTAL GANADO
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 17).FormulaR1C1 = "=SUM(RC[-7]:RC[-1])"

                '########################################################################
                '#############    OTROS DESCUENTOS
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 20).Value = OTROS_DESCUENTOS

                '########################################################################
                '#############    SECCIÓN TOTAL DESCUENTOS
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 21).FormulaR1C1 = "=RC[-3]+RC[-2]+RC[-1]"

                '########################################################################
                '#############    SECCIÓN LIQUIDO PAGABLE
                '########################################################################

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 22).FormulaR1C1 = "=RC[-5]-RC[-1]"

            Next


            colocarRegillas()
            .Cells.EntireColumn.AutoFit()
            .Cells(1, 1).Select

            '########################################################################
            '#############    SECCIÓN APF
            '########################################################################

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            hojaPrePlanilla.Activate()
            .Range(.Cells(2, 2), .Cells(n, 2)).Select

            For Each Celda In .Selection

                hojaResumenAFP.Activate()
                plantillaResumenAFP()

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, -1).Address)
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 1).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, 0).Address)
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 2).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, 1).Address)
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 3).Value = resultadoUb(hojaPrePlanilla.Name, Celda.Offset(0, 16).Address)
                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(0, 4).FormulaR1C1 = "=DATEDIF(" & hojaPrePlanilla.Name & "!R[-2]C,TODAY(),""y"")"

            Next

            hojaANS()
            .Range(.Cells(1, 1), .Cells(1, 2)).CurrentRegion.Select
            formatoTablas()

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            '.Range(.Cells(4, 6), .Cells(n, 15)).Select
            'With .Range(.Selection, .Selection.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            With .Range(.Cells(4, 6), .Cells(n, 15))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlRight
                '    '.Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
                .NumberFormat = "#,##0.00"
            End With

            '.Range(.Cells(4, 1), .Cells(4, 2)).Select
            'With .Range(.Selection, .Selection.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            '    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            '    '.Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            'End With

            '.Range(.Cells(4, 5), .Cells(4, 5)).Select
            'With .Range(.Selection, .Selection.End(Microsoft.Office.Interop.Excel.XlDirection.xlDown))
            '    .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            '    .Font.ColorIndex = Microsoft.Office.Interop.Excel.Constants.xlAutomatic
            'End With


            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(4, 15), .Cells(n, 15)).Select

            For Each Celda In .Selection
                hojaPrePlanilla.Activate()
                .Cells(.Rows.Count, 19).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = resultadoUb(hojaResumenAFP.Name, Celda.Address)
            Next

            '################################################################
            '#############    VARIABLES DATOS ESPECIFICOS DE PRE PLANILLA
            '################################################################

            Dim CI_PRE_PLANILLA As String
            Dim TOTAL_INGRESO_NETO_RUTA As String

            '########################################################################
            '#############    SECCIÓN RC-IVA
            '########################################################################

            hojaPrincipalDatos.Activate()

            SALARIO_MINIMO_NACINAL_VIGENTE_02 = resultadoUb2(hojaPrincipalDatos.Name, .Cells(4, 8).Offset(7, 0).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1))

            hojaPrePlanilla.Activate()
            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            .Range(.Cells(2, 2), .Cells(n, 2)).Select

            For Each Celda In .Selection

                hojaPrePlanilla.Activate()

                CI_PRE_PLANILLA = resultadoUb(hojaPrePlanilla.Name, Celda.Address)
                TOTAL_INGRESO_NETO_RUTA = resultadoUb2(hojaPrePlanilla.Name, Celda.Offset(0, 16).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1)) & "-" & resultadoUb2(hojaPrePlanilla.Name, Celda.Offset(0, 17).Address(ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1))

                hojaResumenRC_IVA.Activate()
                plantillaResumenPlnTrib()

                hojaResumenRC_IVA.Activate()
                .Cells(.Rows.Count, 9).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = CI_PRE_PLANILLA
                .Cells(.Rows.Count, 12).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=IF(ROUND(" & TOTAL_INGRESO_NETO_RUTA & ",0)>(4*" & SALARIO_MINIMO_NACINAL_VIGENTE_02 &
                                                                                "),ROUND(" & TOTAL_INGRESO_NETO_RUTA & ",0),0)"
            Next

            CalculosPlnTrib(hojaPrincipalDatos, hojaResumenRC_IVA)
            recogerResRC_IVA(hojaPrePlanilla, hojaResumenRC_IVA)

            '################################################################
            '#############    APORTE PATRONAL
            '################################################################

            hojaAportePatronal.Activate()
            plantillaResumenAportePatronal()
            CalculosAportePatronal(hojaPrePlanilla, hojaAportePatronal)
            .Cells.EntireColumn.AutoFit()


            '################################################################
            '#############    TEGRESAR HOJA PREPLANILLA
            '################################################################

            hojaPrePlanilla.Activate()
            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            .Range(.Cells(1, 18), .Cells(n, 18)).Font.Bold = True
            .Range(.Cells(1, 18), .Cells(n, 18)).Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Range(.Cells(1, 18), .Cells(n, 18)).Interior.TintAndShade = -0.149998474074526

            .Range(.Cells(1, 22), .Cells(n, 23)).Font.Bold = True
            .Range(.Cells(1, 22), .Cells(n, 23)).Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Range(.Cells(1, 22), .Cells(n, 23)).Interior.TintAndShade = -0.149998474074526

            .Columns("K:W").NumberFormat = "#,##0.00"
            .Cells.EntireColumn.AutoFit()
        End With
    End Sub

    Public Sub generaHojaPlanillaPrePrincipal()
        With Globals.ThisAddIn.Application
            With .Rows("1:1")
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .VerticalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .RowHeight = 45
                .WrapText = True
                .Font.Size = 8
            End With

            .Cells(1, 1).Value = "Nº"
            .Cells(1, 1).Font.Bold = True
            .Cells(1, 2).Value = "Documento de identidad"
            .Cells(1, 2).Font.Bold = True
            .Cells(1, 3).Value = "Apellidos y nombres"
            .Cells(1, 3).Font.Bold = True
            .Cells(1, 4).Value = "País de nacionalidad"
            .Cells(1, 4).Font.Bold = True
            .Cells(1, 5).Value = "Fecha de nacimiento"
            .Cells(1, 5).Font.Bold = True
            .Cells(1, 6).Value = "Sexo (V/M)"
            .Cells(1, 6).Font.Bold = True
            .Cells(1, 7).Value = "Ocupación que desempeña"
            .Cells(1, 7).Font.Bold = True
            .Cells(1, 8).Value = "Fecha de ingreso"
            .Cells(1, 8).Font.Bold = True


            .Cells(1, 9).Value = "Horas pagadas (Día)"
            .Cells(1, 10).Value = "Días pagados (Mes)"
            .Cells(1, 11).Value = "(1) Haber básico"
            .Cells(1, 12).Value = "(2) Bono de Antigüedad"
            .Cells(1, 13).Value = "(3) Bono de producción"
            .Cells(1, 14).Value = "(4) Subsidio de frontera"
            .Cells(1, 15).Value = "(5) Trabajo extraordi-nario y nocturno"
            .Cells(1, 16).Value = "(6) Pago dominical y domingo trabajado"
            .Cells(1, 17).Value = "(7) Otros bonos"

            .Cells(1, 18).Value = "(8) TOTAL GANADO Suma (1 a 7)"
            .Cells(1, 18).Font.Bold = True
            .Cells(1, 18).Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Cells(1, 18).Interior.TintAndShade = -0.149998474074526

            .Cells(1, 19).Value = "(9) Aporte a las AFPs"
            .Cells(1, 20).Value = "(10) RC-IVA"
            .Cells(1, 21).Value = "(11) Otros descuentos"

            .Cells(1, 22).Value = "(12) TOTAL DESCUENTOS Suma (9 a 11)"
            .Cells(1, 22).Font.Bold = True
            .Cells(1, 22).Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Cells(1, 22).Interior.TintAndShade = -0.149998474074526

            .Cells(1, 23).Value = "(13) LÍQUIDO PAGABLE (12-8)"
            .Cells(1, 23).Font.Bold = True
            .Cells(1, 23).Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Cells(1, 23).Interior.TintAndShade = -0.149998474074526

            .Cells(1, 24).Value = "(14) Firma"
            .Cells(1, 24).Font.Bold = True


            .Columns("A:A").ColumnWidth = 2.29
            .Columns("B:B").ColumnWidth = 16
            .Columns("C:C").ColumnWidth = 38
            .Columns("D:D").ColumnWidth = 12.57
            .Columns("E:E").ColumnWidth = 9
            .Columns("F:F").ColumnWidth = 4.29
            .Columns("G:G").ColumnWidth = 15
            .Columns("H:H").ColumnWidth = 9
            .Columns("I:I").ColumnWidth = 6
            .Columns("J:J").ColumnWidth = 6
            .Columns("K:K").ColumnWidth = 7
            .Columns("L:L").ColumnWidth = 9
            .Columns("M:M").ColumnWidth = 7
            .Columns("N:N").ColumnWidth = 7
            .Columns("O:O").ColumnWidth = 7
            .Columns("P:P").ColumnWidth = 7
            .Columns("Q:Q").ColumnWidth = 7
            .Columns("R:R").ColumnWidth = 10
            .Columns("S:S").ColumnWidth = 7
            .Columns("T:T").ColumnWidth = 7
            .Columns("U:U").ColumnWidth = 7
            .Columns("V:V").ColumnWidth = 10
            .Columns("W:W").ColumnWidth = 10
            .Columns("X:X").ColumnWidth = 12
        End With
    End Sub

    Public Sub colocarRegillas()
        With Globals.ThisAddIn.Application
            .Cells(1, 1).CurrentRegion.Select

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

    Public Function resultadoUb(ByVal nombreHojaUb As String, ByVal celdaUb As String) As String

        resultadoUb = "=" & nombreHojaUb & "!" & celdaUb

    End Function

    Public Function resultadoUb2(ByVal nombreHojaUb As String, ByVal celdaUb As String) As String

        resultadoUb2 = nombreHojaUb & "!" & celdaUb

    End Function

End Module
