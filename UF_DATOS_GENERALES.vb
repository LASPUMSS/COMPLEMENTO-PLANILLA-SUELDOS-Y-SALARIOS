Public Class UF_DATOS_GENERALES
    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click

        With Globals.ThisAddIn.Application
            If .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS" And
                .Cells(2, 1).Value = "DATOS GENERALES" And
                .Cells(15, 1).Value = "DATOS ESPECIFICOS" Then

                .Cells(4, 8).Value = txt_NombreRazonSocial.Text
                .Cells(5, 8).Value = txt_NIT.Text
                .Cells(6, 8).Value = txt_NumMisTrb.Text
                .Cells(7, 8).Value = txt_NumCajaSalud.Text
                .Cells(8, 8).Value = txt_GestionPlanilla.Text

                .Cells(9, 8).Value = txt_MesPlanilla.Text
                '.Cells(10, 8).Value = txt_DiaPlanilla.Text
                .Cells(11, 8).Value = txt_SMN.Text
                .Cells(12, 8).Value = txt_UfvIncial.Text
                .Cells(13, 8).Value = txt_UfvFinal.Text

                UF_DATOS_GENERALES.ActiveForm.Close()
            Else
                MsgBox("LA HOJA ACTUAL NO ES LA ADECUADA PARA INTRODUCIR LOS DATOS. GENERE UN HOJA PRINCIPAL PARA EJECUTAR ESTE PROCEDIMIENTO.", MsgBoxStyle.Exclamation)
            End If

        End With



    End Sub

    Private Sub UF_DATOS_GENERALES_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        With Globals.ThisAddIn.Application

            .Cells(4, 8).Value = ""
            .Cells(5, 8).Value = ""
            .Cells(6, 8).Value = ""
            .Cells(7, 8).Value = ""
            .Cells(8, 8).Value = ""

            .Cells(9, 8).Value = ""
            .Cells(10, 8).Value = ""
            .Cells(11, 8).Value = ""
            .Cells(12, 8).Value = ""
            .Cells(13, 8).Value = ""

        End With
    End Sub
End Class