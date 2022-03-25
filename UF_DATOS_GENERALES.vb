Public Class UF_DATOS_GENERALES
    Private Sub btn_Aceptar_Click(sender As Object, e As EventArgs) Handles btn_Aceptar.Click

        With Globals.ThisAddIn.Application

            .Cells(4, 8).Value = txt_NombreRazonSocial.Text
            .Cells(5, 8).Value = txt_NIT.Text
            .Cells(6, 8).Value = txt_NumMisTrb.Text
            .Cells(7, 8).Value = txt_NumCajaSalud.Text
            .Cells(8, 8).Value = txt_GestionPlanilla.Text

            .Cells(9, 8).Value = txt_MesPlanilla.Text
            .Cells(10, 8).Value = NumeroDiasMes(CInt(.Cells(8, 8).Value), CInt(.Cells(9, 8).Value))
            .Cells(11, 8).Value = txt_SMN.Text
            .Cells(12, 8).Value = txt_UfvIncial.Text
            .Cells(13, 8).Value = txt_UfvFinal.Text

            UF_DATOS_GENERALES.ActiveForm.Close()

        End With

    End Sub
    Private Sub UF_DATOS_GENERALES_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        With Globals.ThisAddIn.Application

            txt_NombreRazonSocial.Text = .Cells(4, 8).Value
            txt_NIT.Text = .Cells(5, 8).Value
            txt_NumMisTrb.Text = .Cells(6, 8).Value
            txt_NumCajaSalud.Text = .Cells(7, 8).Value
            txt_GestionPlanilla.Text = .Cells(8, 8).Value

            txt_MesPlanilla.Text = .Cells(9, 8).Value
            '.Cells(10, 8).Value
            txt_SMN.Text = .Cells(11, 8).Value
            txt_UfvIncial.Text = .Cells(12, 8).Value
            txt_UfvFinal.Text = .Cells(13, 8).Value

        End With
    End Sub
End Class