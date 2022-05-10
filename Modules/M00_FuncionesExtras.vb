Module M00_FuncionesExtras
    Public Function NumeroDiasMes(ByVal gestionPlanilla As Long, ByVal mesPlanilla As Integer) As Integer

        '###################################################################################################
        '########    ESTA FUNCION DEVUELVE EL NUMERO DE DIAS QUE TIENE UN MES RESPECTO AL MES Y EL AÑO
        '###################################################################################################

        Dim i2 As Integer
        Dim diasMesGestionNor(11) As Integer
        Dim diasMesGestionBi(11) As Integer
        Dim diasMes As Integer

        diasMesGestionNor(0) = 31
        diasMesGestionNor(1) = 28
        diasMesGestionNor(2) = 31
        diasMesGestionNor(3) = 30
        diasMesGestionNor(4) = 31
        diasMesGestionNor(5) = 30
        diasMesGestionNor(6) = 31
        diasMesGestionNor(7) = 31
        diasMesGestionNor(8) = 30
        diasMesGestionNor(9) = 31
        diasMesGestionNor(10) = 30
        diasMesGestionNor(11) = 31

        diasMesGestionBi(0) = 31
        diasMesGestionBi(1) = 29
        diasMesGestionBi(2) = 31
        diasMesGestionBi(3) = 30
        diasMesGestionBi(4) = 31
        diasMesGestionBi(5) = 30
        diasMesGestionBi(6) = 31
        diasMesGestionBi(7) = 31
        diasMesGestionBi(8) = 30
        diasMesGestionBi(9) = 31
        diasMesGestionBi(10) = 30
        diasMesGestionBi(11) = 31


        If (gestionPlanilla Mod 4) = 0 Then
            'AÑO BICIESTO
            For i2 = 1 To 12
                If mesPlanilla = i2 Then

                    diasMes = diasMesGestionBi(mesPlanilla - 1)

                End If

            Next

        Else
            'AÑO NO BICIESTO
            For i2 = 1 To 12
                If mesPlanilla = i2 Then

                    diasMes = diasMesGestionNor(mesPlanilla - 1)

                End If
            Next
        End If

        Return diasMes
    End Function

    Public Function mesTexto(ByVal mesPlanilla As Integer) As String

        '###################################################################################################
        '########    DEVUELVE EL MES COMO TEXTO PASANDOLOE EL NUMERO DE MES 
        '###################################################################################################

        Dim mesCorrespondiente As String
        mesCorrespondiente = "XXXX"

        If mesPlanilla = 1 Then
            mesCorrespondiente = "ENERO"
        ElseIf mesPlanilla = 2 Then
            mesCorrespondiente = "FEBRERO"
        ElseIf mesPlanilla = 3 Then
            mesCorrespondiente = "MARZO"
        ElseIf mesPlanilla = 4 Then
            mesCorrespondiente = "ABRIL"
        ElseIf mesPlanilla = 5 Then
            mesCorrespondiente = "MAYO"
        ElseIf mesPlanilla = 6 Then
            mesCorrespondiente = "JUNIO"
        ElseIf mesPlanilla = 7 Then
            mesCorrespondiente = "JULIO"
        ElseIf mesPlanilla = 8 Then
            mesCorrespondiente = "AGOSTO"
        ElseIf mesPlanilla = 9 Then
            mesCorrespondiente = "SEPTIEMBRE"
        ElseIf mesPlanilla = 10 Then
            mesCorrespondiente = "OCTUBRE"
        ElseIf mesPlanilla = 11 Then
            mesCorrespondiente = "NOVIEMBRE"
        ElseIf mesPlanilla = 12 Then
            mesCorrespondiente = "DICIEMBRE"
        Else
            mesCorrespondiente = "XXXX"
        End If

        Return mesCorrespondiente

    End Function
End Module
