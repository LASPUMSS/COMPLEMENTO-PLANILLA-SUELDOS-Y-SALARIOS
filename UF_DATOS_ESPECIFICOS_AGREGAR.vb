Public Class UF_DATOS_ESPECIFICOS_AGREGAR
    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click
        With Globals.ThisAddIn.Application

            If txt_NombreTrabajador.Text <> "" Then

                .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = .Cells(.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Row - 16
                .Cells(.Rows.Count, 2).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_CI.Text
                .Cells(.Rows.Count, 3).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=RC[24] & "" ""  & RC[25] & "" "" & RC[23]"
                .Cells(.Rows.Count, 4).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_PaisNacionalidad.Text
                .Cells(.Rows.Count, 5).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_GestionNac.Text & "/" & txt_MesNac.Text & "/" & txt_DiaNac.Text
                .Cells(.Rows.Count, 6).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_GestionIng.Text & "/" & txt_MesIng.Text & "/" & txt_DiaIng.Text
                .Cells(.Rows.Count, 7).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_Sexo.Text
                .Cells(.Rows.Count, 8).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_OcuapacionQueDesp.Text
                .Cells(.Rows.Count, 9).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_HorasPagadas.Text
                .Cells(.Rows.Count, 10).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_DiasPagados.Text
                .Cells(.Rows.Count, 11).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_HaberBasico.Text
                .Cells(.Rows.Count, 12).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_BonoProduccion.Text
                .Cells(.Rows.Count, 13).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=IF(RC[1]=""CORRESPONDE"",RC[-2]*25%,0)"
                .Cells(.Rows.Count, 14).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_CorrespondeBonoFronteras.Text
                .Cells(.Rows.Count, 15).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_HorasExtraordinarias.Text
                .Cells(.Rows.Count, 16).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_HorasNocturnas.Text
                .Cells(.Rows.Count, 17).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_CategoriaTrabajoNoc.Text
                .Cells(.Rows.Count, 18).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_HorasDomingos.Text
                .Cells(.Rows.Count, 19).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).FormulaR1C1 = "=IF(AND(RC[1]=""OBRERO"",RC[2]=""LLEGO PUNTUAL""),""CUMPLE"",""NO CUMPLE"")"
                .Cells(.Rows.Count, 20).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_DominicalOcupacion.Text
                .Cells(.Rows.Count, 21).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_LlegoPuntual.Text
                .Cells(.Rows.Count, 22).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_OtrosBonos.Text
                .Cells(.Rows.Count, 23).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_ConceptoOtrosBonos.Text
                .Cells(.Rows.Count, 24).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_OtrosDescuentos.Text
                .Cells(.Rows.Count, 25).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_ConceptoOtrosDescuentos.Text
                .Cells(.Rows.Count, 26).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_NombreTrabajador.Text
                .Cells(.Rows.Count, 27).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_ApellidoPat.Text
                .Cells(.Rows.Count, 28).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_ApellidoMat.Text
                .Cells(.Rows.Count, 29).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_CodigoDependienteRCIVA.Text
                .Cells(.Rows.Count, 30).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_TipoDocumentoRCIVA.Text
                .Cells(.Rows.Count, 31).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = ComB_NovedadesRCIVA.Text
                .Cells(.Rows.Count, 32).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_Form110.Text
                .Cells(.Rows.Count, 33).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Offset(1, 0).Value = txt_SaldoRcIvaMesAnt.Text

                .Cells.EntireColumn.AutoFit()

                UF_DATOS_ESPECIFICOS_AGREGAR.ActiveForm.Close()

            Else
                MsgBox("DEBE COMPLETAR TODOS LOS CAMPOS DEL FORMULARIO", MsgBoxStyle.Exclamation)
            End If
        End With
    End Sub

    Private Sub UF_DATOS_ESPECIFICOS_AGREGAR_Activated(sender As Object, e As EventArgs) Handles Me.Activated

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