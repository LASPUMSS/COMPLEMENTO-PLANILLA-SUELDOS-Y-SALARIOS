Module M01_GenerarHojaPrincipalPlanilla
    Public Sub generarPlantillaHojaPrincipal()
        With Globals.ThisAddIn.Application

            .Sheets.Add()
            .Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter

            .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS"
            With .Range(.Cells(1, 1), .Cells(1, 33))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Font.Size = 72
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With

            .Cells(2, 1).Value = "DATOS GENERALES"
            With .Range(.Cells(2, 1), .Cells(2, 33))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Font.Size = 36
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
            End With

            .Range(.Cells(4, 1), .Cells(4, 6)).Merge()
            .Range(.Cells(5, 1), .Cells(5, 6)).Merge()
            .Range(.Cells(6, 1), .Cells(6, 6)).Merge()
            .Range(.Cells(7, 1), .Cells(7, 6)).Merge()

            .Range(.Cells(8, 1), .Cells(8, 6)).Merge()
            .Range(.Cells(9, 1), .Cells(9, 6)).Merge()
            .Range(.Cells(10, 1), .Cells(10, 6)).Merge()
            .Range(.Cells(11, 1), .Cells(11, 6)).Merge()
            .Range(.Cells(12, 1), .Cells(12, 6)).Merge()
            .Range(.Cells(13, 1), .Cells(13, 6)).Merge()


            .Range(.Cells(4, 1), .Cells(4, 1)).Value = "NOMBRE O RAZON SOCIAL:"
            .Range(.Cells(5, 1), .Cells(5, 1)).Value = "NUMERO DE IDENTIFICACIÓN TRIBUTARIA (NIT):"
            .Range(.Cells(6, 1), .Cells(6, 1)).Value = "NUMERO IDENTIFICADOR DEL EMPLEADOR ANTE EL MINISTERIO DE TRABAJO:"
            .Range(.Cells(7, 1), .Cells(7, 1)).Value = "NUMERO DE EMPLEADOR (CAJA DE SALUD):"

            .Range(.Cells(8, 1), .Cells(8, 1)).Value = "AÑO PLANILLA:"
            .Range(.Cells(9, 1), .Cells(9, 1)).Value = "MES PLANILLA:"
            .Range(.Cells(10, 1), .Cells(10, 1)).Value = "DIA PLANILLA:"
            .Range(.Cells(11, 1), .Cells(11, 1)).Value = "SALARIO MINIMO NACIONAL (VIGENTE):"
            .Range(.Cells(12, 1), .Cells(12, 1)).Value = "UFV DEL ULTIMO DIA HABIL DEL MES A DECLARAR:"
            .Range(.Cells(13, 1), .Cells(13, 1)).Value = "UFV DEL ULTIMO DIA HABIL DEL MES ANTERIOR A DECLARAR:"

            With .Range(.Cells(4, 1), .Cells(13, 6))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Size = 18
                .Font.Bold = True
            End With

            .Cells(15, 1).Value = "DATOS ESPECIFICOS"
            With .Range(.Cells(15, 1), .Cells(15, 33))
                .Merge()
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlLeft
                .Font.Size = 36
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = 0.399975585192419
            End With

            .Cells(16, 1).Value = "Nº"
            .Cells(16, 2).Value = "DOCUMENTO DE IDENTIDAD"
            .Cells(16, 3).Value = "APELLIDO Y NOMBRES"
            .Cells(16, 4).Value = "PAIS DE NACIONALIDAD"
            .Cells(16, 5).Value = "FECHA DE NACIMIENTO"
            .Cells(16, 6).Value = "FECHA DE INGRESO"
            .Cells(16, 7).Value = "SEXO(V/M)"
            .Cells(16, 8).Value = "OCUPACIÓN QUE DESEMPEÑA"
            .Cells(16, 9).Value = "HORAS PAGADAS (DIA)"
            .Cells(16, 10).Value = "DIAS PAGADAS (MES)"
            .Cells(16, 11).Value = "HABER BASICO"
            .Cells(16, 12).Value = "BONO PRODUCCIÓN"
            .Cells(16, 13).Value = "SUBSIDIO FRONTERAS"
            .Cells(16, 14).Value = "CONDICIÓN SUBSIDIO FRONTERAS"
            .Cells(16, 15).Value = "HORAS EXTRAORDINARIAS TRABAJADAS"
            .Cells(16, 16).Value = "HORAS NOCTURNAS TRABAJADAS"
            .Cells(16, 17).Value = "CATEGORIA DE TRABAJO NOCTURNO"
            .Cells(16, 18).Value = "HORAS TRABAJADOS EN DOMINGO"
            .Cells(16, 19).Value = "DOMINICAL"
            .Cells(16, 20).Value = "CONDICIÓN 1 DOMINICAL"
            .Cells(16, 21).Value = "CONDICIÓN 2 DOMINICAL"
            .Cells(16, 22).Value = "OTROS BONOS"
            .Cells(16, 23).Value = "CONCEPTO DE OTROS BONOS"
            .Cells(16, 24).Value = "OTROS DESCUENTOS"
            .Cells(16, 25).Value = "CONCEPTO DE OTROS DESCUENTOS"

            .Cells(16, 26).Value = "NOMBRE(S)"
            .Cells(16, 27).Value = "APELLIDO PATERNO"
            .Cells(16, 28).Value = "APELLIDO MATERNO"
            .Cells(16, 29).Value = "CODIGO DEPENDIENTE RC-IVA"
            .Cells(16, 30).Value = "TIPO DE DOCUMENTO"
            .Cells(16, 31).Value = "NOVEDADES I - V - D"
            .Cells(16, 32).Value = "FORM 110"
            .Cells(16, 33).Value = "SALDO A FAVOR DE DEPENDIENTE DEL PERIODO ANTERIO"

            With .Range(.Cells(16, 1), .Cells(16, 33))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .Font.Bold = True
                .Font.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Pattern = Microsoft.Office.Interop.Excel.Constants.xlSolid
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5
                .Interior.TintAndShade = -0.249977111117893
            End With

            .Cells(4, 8).HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
            With .Range(.Cells(5, 8), .Cells(11, 8))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "0"
            End With

            With .Range(.Cells(12, 8), .Cells(13, 8))
                .HorizontalAlignment = Microsoft.Office.Interop.Excel.Constants.xlCenter
                .NumberFormat = "0.00000000"
            End With

            .ActiveWindow.Zoom = 60
            .Cells.EntireColumn.AutoFit()
        End With
    End Sub

End Module
