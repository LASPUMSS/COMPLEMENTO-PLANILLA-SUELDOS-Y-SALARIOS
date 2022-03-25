Module M00_FuncionesExtras
    Public Function NumeroDiasMes(ByVal gestionPlanilla As Long, ByVal mesPlanilla As Integer) As Integer

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
End Module
