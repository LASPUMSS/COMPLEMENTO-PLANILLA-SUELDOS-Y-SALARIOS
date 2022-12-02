﻿Public Class ThisAddIn
    Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
        Return New Ribbon1()
    End Function

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ' PERMITIR QUE EL LIBRO SE PUEDA EDITAR
        Globals.ThisAddIn.Application.ActiveWorkbook.ChangeFileAccess(Microsoft.Office.Interop.Excel.XlFileAccess.xlReadWrite)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
