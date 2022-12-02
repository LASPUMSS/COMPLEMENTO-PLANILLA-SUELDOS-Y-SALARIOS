'TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

'1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

'Protected Overrides Function CreateRibbonExtensibilityObject() As Microsoft.Office.Core.IRibbonExtensibility
'    Return New Ribbon1()
'End Function

'2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
'   como hacer clic en un botón. Nota: si ha exportado esta cinta desde el
'   diseñador de la cinta de opciones, deberá mover el código de los controladores de eventos a los métodos de devolución de llamada y
'   modificar el código para que funcione con el modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

'3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.

'Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.

<Runtime.InteropServices.ComVisible(True)> _
Public Class Ribbon1
    Implements Office.IRibbonExtensibility

    Private ribbon As Office.IRibbonUI

    Public Sub New()
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements Office.IRibbonExtensibility.GetCustomUI
        Return GetResourceText("COMPLEMENTO_PLANILLA_SUELDOS_Y_SALARIOS.Ribbon1.xml")
    End Function

    Public Sub btn_GenerarHojaPrincipalDatos(Control As Office.IRibbonControl)
        generarPlantillaHojaPrincipal()
    End Sub

    Public Sub btn_AgregarDatosGenerales(Control As Office.IRibbonControl)
        Dim frm As New UF_DATOS_GENERALES

        With Globals.ThisAddIn.Application
            If .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS" And
                .Cells(2, 1).Value = "DATOS GENERALES" And
                .Cells(15, 1).Value = "DATOS ESPECIFICOS" Then

                frm.ShowDialog()
            Else
                MsgBox("LA HOJA ACTUAL NO ES LA ADECUADA PARA INTRODUCIR LOS DATOS. GENERE UN HOJA PRINCIPAL PARA EJECUTAR ESTE PROCEDIMIENTO.", MsgBoxStyle.Exclamation)
            End If
        End With

    End Sub
    Public Sub btn_AgregarDatosEspecificos(Control As Office.IRibbonControl)
        Dim frm As New UF_DATOS_ESPECIFICOS_AGREGAR

        With Globals.ThisAddIn.Application
            If .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS" And
                .Cells(2, 1).Value = "DATOS GENERALES" And
                .Cells(15, 1).Value = "DATOS ESPECIFICOS" Then

                frm.ShowDialog()
            Else
                MsgBox("LA HOJA ACTUAL NO ES LA ADECUADA PARA INTRODUCIR LOS DATOS. GENERE UN HOJA PRINCIPAL PARA EJECUTAR ESTE PROCEDIMIENTO.", MsgBoxStyle.Exclamation)
            End If
        End With
    End Sub
    Public Sub btn_ModificarDatosEspecificos(Control As Office.IRibbonControl)
        Dim frm As New UF_DATOS_ESPECIFICOS_MODIFICAR

        With Globals.ThisAddIn.Application
            If .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS" And
                .Cells(2, 1).Value = "DATOS GENERALES" And
                .Cells(15, 1).Value = "DATOS ESPECIFICOS" Then

                frm.ShowDialog()
            Else
                MsgBox("LA HOJA ACTUAL NO ES LA ADECUADA PARA INTRODUCIR LOS DATOS. GENERE UN HOJA PRINCIPAL PARA EJECUTAR ESTE PROCEDIMIENTO.", MsgBoxStyle.Exclamation)
            End If
        End With
    End Sub
    Public Sub btn_GenerarPlanillaSueldosSalarios(Control As Office.IRibbonControl)

        With Globals.ThisAddIn.Application
            If .Cells(1, 1).Value = "ELABORACIÓN DE PLANILLA DE SUELDOS Y SALARIOS" And
                .Cells(2, 1).Value = "DATOS GENERALES" And
                .Cells(15, 1).Value = "DATOS ESPECIFICOS" Then

                hojaPrinciplaPlanillaSueldosSalarios()
            Else
                MsgBox("LA HOJA ACTUAL NO ES LA ADECUADA PARA INTRODUCIR LOS DATOS. GENERE UN HOJA PRINCIPAL PARA EJECUTAR ESTE PROCEDIMIENTO.", MsgBoxStyle.Exclamation)
            End If
        End With
    End Sub

    Public Sub btn_GenerarDatosEjemplo(Control As Office.IRibbonControl)
        datosDeEjemploDeSueldos()
    End Sub
#Region "Devoluciones de llamada de la cinta de opciones"
    'Cree métodos de devolución de llamada aquí. Para obtener más información sobre la adición de métodos de devolución de llamada, visite https://go.microsoft.com/fwlink/?LinkID=271226
    Public Sub Ribbon_Load(ByVal ribbonUI As Office.IRibbonUI)
        Me.ribbon = ribbonUI
    End Sub



#End Region

#Region "Asistentes"

    Private Shared Function GetResourceText(ByVal resourceName As String) As String
        Dim asm As Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly()
        Dim resourceNames() As String = asm.GetManifestResourceNames()
        For i As Integer = 0 To resourceNames.Length - 1
            If String.Compare(resourceName, resourceNames(i), StringComparison.OrdinalIgnoreCase) = 0 Then
                Using resourceReader As IO.StreamReader = New IO.StreamReader(asm.GetManifestResourceStream(resourceNames(i)))
                    If resourceReader IsNot Nothing Then
                        Return resourceReader.ReadToEnd()
                    End If
                End Using
            End If
        Next
        Return Nothing
    End Function

#End Region

End Class
