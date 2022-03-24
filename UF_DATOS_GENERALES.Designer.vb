<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UF_DATOS_GENERALES
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txt_NombreRazonSocial = New System.Windows.Forms.TextBox()
        Me.txt_NIT = New System.Windows.Forms.TextBox()
        Me.txt_NumMisTrb = New System.Windows.Forms.TextBox()
        Me.txt_NumCajaSalud = New System.Windows.Forms.TextBox()
        Me.txt_UfvFina = New System.Windows.Forms.TextBox()
        Me.txt_UfvIncial = New System.Windows.Forms.TextBox()
        Me.txt_SMN = New System.Windows.Forms.TextBox()
        Me.txt_GestionPlanilla = New System.Windows.Forms.TextBox()
        Me.txt_MesPlanilla = New System.Windows.Forms.TextBox()
        Me.txt_DiaPlanilla = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(150, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "NOMBRE O RAZON SOCIAL:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 88)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(259, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "NUMERO DE IDENTIFICACION TRIBUTARIA (NIT):"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 128)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(422, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "NUMERO IDENTIFICADOR DEL EMPLEADOR ANTE EL MINISTERIO DE TRABAJO:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(26, 168)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(238, 13)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "NUMERO DE EMPLEADOR (CAJA DE SALUD):"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(26, 208)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(334, 13)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "UFV DEL ULTIMO DIA HABIL DEL MES ANTERIOR A DECLARAR:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(26, 248)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(275, 13)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "UFV DEL ULTIMO DIA HABIL DEL MES A DECLARAR:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(26, 288)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(204, 13)
        Me.Label7.TabIndex = 6
        Me.Label7.Text = "SALARIO MINIMO NACINAL (VIGENTE):"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(26, 328)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(86, 13)
        Me.Label8.TabIndex = 7
        Me.Label8.Text = "AÑO PLANILLA:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(26, 368)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(86, 13)
        Me.Label9.TabIndex = 8
        Me.Label9.Text = "MES PLANILLA:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(26, 408)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(81, 13)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "DIA PLANILLA:"
        '
        'txt_NombreRazonSocial
        '
        Me.txt_NombreRazonSocial.Location = New System.Drawing.Point(483, 45)
        Me.txt_NombreRazonSocial.Name = "txt_NombreRazonSocial"
        Me.txt_NombreRazonSocial.Size = New System.Drawing.Size(378, 20)
        Me.txt_NombreRazonSocial.TabIndex = 10
        '
        'txt_NIT
        '
        Me.txt_NIT.Location = New System.Drawing.Point(483, 85)
        Me.txt_NIT.Name = "txt_NIT"
        Me.txt_NIT.Size = New System.Drawing.Size(378, 20)
        Me.txt_NIT.TabIndex = 11
        '
        'txt_NumMisTrb
        '
        Me.txt_NumMisTrb.Location = New System.Drawing.Point(483, 125)
        Me.txt_NumMisTrb.Name = "txt_NumMisTrb"
        Me.txt_NumMisTrb.Size = New System.Drawing.Size(378, 20)
        Me.txt_NumMisTrb.TabIndex = 12
        '
        'txt_NumCajaSalud
        '
        Me.txt_NumCajaSalud.Location = New System.Drawing.Point(483, 165)
        Me.txt_NumCajaSalud.Name = "txt_NumCajaSalud"
        Me.txt_NumCajaSalud.Size = New System.Drawing.Size(378, 20)
        Me.txt_NumCajaSalud.TabIndex = 13
        '
        'txt_UfvFina
        '
        Me.txt_UfvFina.Location = New System.Drawing.Point(483, 205)
        Me.txt_UfvFina.Name = "txt_UfvFina"
        Me.txt_UfvFina.Size = New System.Drawing.Size(378, 20)
        Me.txt_UfvFina.TabIndex = 14
        '
        'txt_UfvIncial
        '
        Me.txt_UfvIncial.Location = New System.Drawing.Point(483, 245)
        Me.txt_UfvIncial.Name = "txt_UfvIncial"
        Me.txt_UfvIncial.Size = New System.Drawing.Size(378, 20)
        Me.txt_UfvIncial.TabIndex = 15
        '
        'txt_SMN
        '
        Me.txt_SMN.Location = New System.Drawing.Point(483, 285)
        Me.txt_SMN.Name = "txt_SMN"
        Me.txt_SMN.Size = New System.Drawing.Size(378, 20)
        Me.txt_SMN.TabIndex = 16
        '
        'txt_GestionPlanilla
        '
        Me.txt_GestionPlanilla.Location = New System.Drawing.Point(483, 325)
        Me.txt_GestionPlanilla.Name = "txt_GestionPlanilla"
        Me.txt_GestionPlanilla.Size = New System.Drawing.Size(378, 20)
        Me.txt_GestionPlanilla.TabIndex = 17
        Me.txt_GestionPlanilla.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_MesPlanilla
        '
        Me.txt_MesPlanilla.Location = New System.Drawing.Point(483, 365)
        Me.txt_MesPlanilla.Name = "txt_MesPlanilla"
        Me.txt_MesPlanilla.Size = New System.Drawing.Size(378, 20)
        Me.txt_MesPlanilla.TabIndex = 18
        Me.txt_MesPlanilla.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txt_DiaPlanilla
        '
        Me.txt_DiaPlanilla.Location = New System.Drawing.Point(483, 405)
        Me.txt_DiaPlanilla.Name = "txt_DiaPlanilla"
        Me.txt_DiaPlanilla.Size = New System.Drawing.Size(378, 20)
        Me.txt_DiaPlanilla.TabIndex = 19
        Me.txt_DiaPlanilla.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(788, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(94, 23)
        Me.Button1.TabIndex = 20
        Me.Button1.Text = "ACEPTAR"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'UF_DATOS_GENERALES
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(891, 450)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txt_DiaPlanilla)
        Me.Controls.Add(Me.txt_MesPlanilla)
        Me.Controls.Add(Me.txt_GestionPlanilla)
        Me.Controls.Add(Me.txt_SMN)
        Me.Controls.Add(Me.txt_UfvIncial)
        Me.Controls.Add(Me.txt_UfvFina)
        Me.Controls.Add(Me.txt_NumCajaSalud)
        Me.Controls.Add(Me.txt_NumMisTrb)
        Me.Controls.Add(Me.txt_NIT)
        Me.Controls.Add(Me.txt_NombreRazonSocial)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "UF_DATOS_GENERALES"
        Me.Text = "DATOS GENERALES PLANILLA"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents txt_NombreRazonSocial As Windows.Forms.TextBox
    Friend WithEvents txt_NIT As Windows.Forms.TextBox
    Friend WithEvents txt_NumMisTrb As Windows.Forms.TextBox
    Friend WithEvents txt_NumCajaSalud As Windows.Forms.TextBox
    Friend WithEvents txt_UfvFina As Windows.Forms.TextBox
    Friend WithEvents txt_UfvIncial As Windows.Forms.TextBox
    Friend WithEvents txt_SMN As Windows.Forms.TextBox
    Friend WithEvents txt_GestionPlanilla As Windows.Forms.TextBox
    Friend WithEvents txt_MesPlanilla As Windows.Forms.TextBox
    Friend WithEvents txt_DiaPlanilla As Windows.Forms.TextBox
    Friend WithEvents Button1 As Windows.Forms.Button
End Class
