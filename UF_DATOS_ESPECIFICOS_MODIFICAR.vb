Public Class UF_DATOS_ESPECIFICOS_MODIFICAR
    Private Sub UF_DATOS_ESPECIFICOS_MODIFICAR_Activated(sender As Object, e As EventArgs) Handles Me.Activated

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Long

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            'MsgBox(n)

            ComB_CI.Items.Clear()

            For i = 17 To n
                ComB_CI.Items.Add(.Cells(i, 2).Value)
            Next

        End With

        ComB_CorrespondeBonoFronteras.Items.Clear()
        ComB_CorrespondeBonoFronteras.Items.Add("CORRESPONDE")
        ComB_CorrespondeBonoFronteras.Items.Add("NO CORRESPONDE")

        ComB_CategoriaTrabajoNoc.Items.Clear()
        ComB_CategoriaTrabajoNoc.Items.Add("RECARGO NOCTURNO")
        ComB_CategoriaTrabajoNoc.Items.Add("HOMBRES ADMINISTRATIVOS")
        ComB_CategoriaTrabajoNoc.Items.Add("OBREROS")
        ComB_CategoriaTrabajoNoc.Items.Add("MUJERES O MENOR 18")
        ComB_CategoriaTrabajoNoc.Items.Add("TRABAJADORES DE ALTO RIESGO")

        ComB_DominicalOcupacion.Items.Clear()
        ComB_DominicalOcupacion.Items.Add("ADMINISTRACION")
        ComB_DominicalOcupacion.Items.Add("OBRERO")

        ComB_LlegoPuntual.Items.Clear()
        ComB_LlegoPuntual.Items.Add("LLEGO PUNTUAL")
        ComB_LlegoPuntual.Items.Add("NO LLEGO PUNTUAL")

        ComB_NovedadesRCIVA.Items.Clear()
        ComB_NovedadesRCIVA.Items.Add("INCORPORACION")
        ComB_NovedadesRCIVA.Items.Add("VIGENTE")
        ComB_NovedadesRCIVA.Items.Add("DESVINCULADO")

        ComB_Sexo.Items.Clear()
        ComB_Sexo.Items.Add("VARON")
        ComB_Sexo.Items.Add("MUJER")

    End Sub

    Private Sub ComB_CI_TextChanged(sender As Object, e As EventArgs) Handles ComB_CI.TextChanged
        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Long
            Dim celda As Excel.Range

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            For i = 17 To n

                celda = .Cells(i, 2)
                If ComB_CI.Text = .Cells(i, 2).Value Then
                    '##############################################################################
                    '########################   DATOS DEL TRABAJADOR
                    '##############################################################################

                    txt_NombreTrabajador.Text = celda.Offset(0, 24).Value
                    txt_ApellidoPat.Text = celda.Offset(0, 25).Value
                    txt_ApellidoMat.Text = celda.Offset(0, 26).Value
                    txt_HaberBasico.Text = celda.Offset(0, 9).Value
                    txt_DiasPagados.Text = celda.Offset(0, 7).Value
                    txt_PaisNacionalidad.Text = celda.Offset(0, 2).Value
                    txt_OcuapacionQueDesp.Text = celda.Offset(0, 6).Value
                    ComB_Sexo.Text = celda.Offset(0, 5).Value
                    txt_HorasPagadas.Text = celda.Offset(0, 7).Value
                    txt_GestionNac.Text = DatePart("yyyy", celda.Offset(0, 3).Value)
                    txt_MesNac.Text = DatePart("m", celda.Offset(0, 3).Value)
                    txt_DiaNac.Text = DatePart("d", celda.Offset(0, 3).Value)
                    txt_GestionIng.Text = DatePart("yyyy", celda.Offset(0, 4).Value)
                    txt_MesIng.Text = DatePart("m", celda.Offset(0, 4).Value)
                    txt_DiaIng.Text = DatePart("d", celda.Offset(0, 4).Value)

                    '##############################################################################
                    '########################   DOMINICAL
                    '##############################################################################

                    ComB_DominicalOcupacion.Text = celda.Offset(0, 18).Value
                    ComB_LlegoPuntual.Text = celda.Offset(0, 19).Value

                    '##############################################################################
                    '########################   DATOS DE HABERES EXTRAS DEL TRABAJADOR
                    '##############################################################################

                    ComB_CorrespondeBonoFronteras.Text = celda.Offset(0, 12).Value
                    txt_BonoProduccion.Text = celda.Offset(0, 10).Value
                    txt_HorasExtraordinarias.Text = celda.Offset(0, 13).Value
                    txt_HorasNocturnas.Text = celda.Offset(0, 14).Value
                    ComB_CategoriaTrabajoNoc.Text = celda.Offset(0, 15).Value
                    txt_HorasDomingos.Text = celda.Offset(0, 16).Value

                    '##############################################################################
                    '########################   DATOS RC-IVA
                    '##############################################################################

                    txt_CodigoDependienteRCIVA.Text = celda.Offset(0, 27).Value
                    txt_TipoDocumentoRCIVA.Text = celda.Offset(0, 28).Value
                    ComB_NovedadesRCIVA.Text = celda.Offset(0, 29).Value
                    txt_Form110.Text = celda.Offset(0, 30).Value
                    txt_SaldoRcIvaMesAnt.Text = celda.Offset(0, 31).Value

                    '##############################################################################
                    '########################   OTROS INGRESOS E EGRESOS DEL TRABAJADOR
                    '##############################################################################

                    txt_OtrosBonos.Text = celda.Offset(0, 20).Value
                    txt_ConceptoOtrosBonos.Text = celda.Offset(0, 21).Value
                    txt_OtrosDescuentos.Text = celda.Offset(0, 22).Value
                    txt_ConceptoOtrosDescuentos.Text = celda.Offset(0, 23).Value


                    Exit For
                End If
            Next
        End With
    End Sub

    Private Sub btn_Modificar_Click(sender As Object, e As EventArgs) Handles btn_Modificar.Click

        With Globals.ThisAddIn.Application
            Dim n As Long
            Dim i As Long
            Dim celda As Excel.Range

            n = .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

            For i = 17 To n

                celda = .Cells(i, 2)
                If ComB_CI.Text = .Cells(i, 2).Value Then

                    celda.Offset(0, -1).Value = celda.Row - 16

                    celda.Value = ComB_CI.Text

                    celda.Offset(0, 1).FormulaR1C1 = "=RC[24] & "" ""  & RC[25] & "" "" & RC[23]"
                    celda.Offset(0, 2).Value = txt_PaisNacionalidad.Text
                    celda.Offset(0, 3).Value = txt_GestionNac.Text & "/" & txt_MesNac.Text & "/" & txt_DiaNac.Text
                    celda.Offset(0, 4).Value = txt_GestionIng.Text & "/" & txt_MesIng.Text & "/" & txt_DiaIng.Text
                    celda.Offset(0, 5).Value = ComB_Sexo.Text
                    celda.Offset(0, 6).Value = txt_OcuapacionQueDesp.Text
                    celda.Offset(0, 7).Value = txt_HorasPagadas.Text
                    celda.Offset(0, 8).Value = txt_DiasPagados.Text
                    celda.Offset(0, 9).Value = txt_HaberBasico.Text
                    celda.Offset(0, 10).Value = txt_BonoProduccion.Text
                    celda.Offset(0, 11).FormulaR1C1 = "=IF(RC[1]=""CORRESPONDE"",RC[-2]*25%,0)"
                    celda.Offset(0, 12).Value = ComB_CorrespondeBonoFronteras.Text
                    celda.Offset(0, 13).Value = txt_HorasExtraordinarias.Text
                    celda.Offset(0, 14).Value = txt_HorasNocturnas.Text
                    celda.Offset(0, 15).Value = ComB_CategoriaTrabajoNoc.Text
                    celda.Offset(0, 16).Value = txt_HorasDomingos.Text
                    celda.Offset(0, 17).FormulaR1C1 = "=IF(AND(RC[1]=""OBRERO"",RC[2]=""LLEGO PUNTUAL""),""CUMPLE"",""NO CUMPLE"")"
                    celda.Offset(0, 18).Value = ComB_DominicalOcupacion.Text
                    celda.Offset(0, 19).Value = ComB_LlegoPuntual.Text
                    celda.Offset(0, 20).Value = txt_OtrosBonos.Text
                    celda.Offset(0, 21).Value = txt_ConceptoOtrosBonos.Text
                    celda.Offset(0, 22).Value = txt_OtrosDescuentos.Text
                    celda.Offset(0, 23).Value = txt_ConceptoOtrosDescuentos.Text
                    celda.Offset(0, 24).Value = txt_NombreTrabajador.Text
                    celda.Offset(0, 25).Value = txt_ApellidoPat.Text
                    celda.Offset(0, 26).Value = txt_ApellidoMat.Text
                    celda.Offset(0, 27).Value = txt_CodigoDependienteRCIVA.Text
                    celda.Offset(0, 28).Value = txt_TipoDocumentoRCIVA.Text
                    celda.Offset(0, 29).Value = ComB_NovedadesRCIVA.Text
                    celda.Offset(0, 30).Value = txt_Form110.Text
                    celda.Offset(0, 31).Value = txt_SaldoRcIvaMesAnt.Text

                    .Cells.EntireColumn.AutoFit()

                    Exit For
                End If
            Next
        End With
    End Sub

    Private Sub ComB_DominicalOcupacion_TextChanged(sender As Object, e As EventArgs) Handles ComB_DominicalOcupacion.TextChanged
        If ComB_DominicalOcupacion.Text = "OBRERO" And ComB_LlegoPuntual.Text = "LLEGO PUNTUAL" Then
            txt_CorrespondeDominical.ReadOnly = False
            txt_CorrespondeDominical.Text = "CORRESPONDE"
            txt_CorrespondeDominical.ReadOnly = True
        Else
            txt_CorrespondeDominical.ReadOnly = False
            txt_CorrespondeDominical.Text = "NO CORRESPONDE"
            txt_CorrespondeDominical.ReadOnly = True
        End If
    End Sub

    Private Sub ComB_LlegoPuntual_TextChanged(sender As Object, e As EventArgs) Handles ComB_LlegoPuntual.TextChanged
        If ComB_DominicalOcupacion.Text = "OBRERO" And ComB_LlegoPuntual.Text = "LLEGO PUNTUAL" Then
            txt_CorrespondeDominical.ReadOnly = False
            txt_CorrespondeDominical.Text = "CORRESPONDE"
            txt_CorrespondeDominical.ReadOnly = True
        Else
            txt_CorrespondeDominical.ReadOnly = False
            txt_CorrespondeDominical.Text = "NO CORRESPONDE"
            txt_CorrespondeDominical.ReadOnly = True
        End If
    End Sub
End Class