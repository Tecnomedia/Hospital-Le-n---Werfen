<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogoConfiguracion
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.cboTipoLocal = New System.Windows.Forms.ComboBox
        Me.txtcnLocal = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtDSN = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtRegExp = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cboTipoConsulta = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.cboConectando = New System.Windows.Forms.ComboBox
        Me.cboMuestraMicro = New System.Windows.Forms.ComboBox
        Me.lblMuestraMicro = New System.Windows.Forms.Label
        Me.btnProbarConexionLocal = New System.Windows.Forms.Button
        Me.btnProbarConexionRemota = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 90)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Tipo BBDD Local"
        '
        'cboTipoLocal
        '
        Me.cboTipoLocal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipoLocal.FormattingEnabled = True
        Me.cboTipoLocal.Items.AddRange(New Object() {"Access", "SQL Server"})
        Me.cboTipoLocal.Location = New System.Drawing.Point(105, 86)
        Me.cboTipoLocal.Name = "cboTipoLocal"
        Me.cboTipoLocal.Size = New System.Drawing.Size(121, 21)
        Me.cboTipoLocal.TabIndex = 1
        '
        'txtcnLocal
        '
        Me.txtcnLocal.Location = New System.Drawing.Point(105, 113)
        Me.txtcnLocal.Name = "txtcnLocal"
        Me.txtcnLocal.Size = New System.Drawing.Size(470, 20)
        Me.txtcnLocal.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 116)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "ConnectionString"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 63)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(149, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "BASE DE DATOS LOCAL"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(11, 151)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(162, 13)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "BASE DE DATOS REMOTA"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(11, 178)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 13)
        Me.Label5.TabIndex = 6
        Me.Label5.Text = "DSN"
        '
        'txtDSN
        '
        Me.txtDSN.Location = New System.Drawing.Point(105, 175)
        Me.txtDSN.Name = "txtDSN"
        Me.txtDSN.Size = New System.Drawing.Size(121, 20)
        Me.txtDSN.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(11, 218)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(85, 13)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "MISCELANEA"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(11, 245)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 13)
        Me.Label7.TabIndex = 9
        Me.Label7.Text = "Reg Exp"
        '
        'txtRegExp
        '
        Me.txtRegExp.Location = New System.Drawing.Point(105, 242)
        Me.txtRegExp.Name = "txtRegExp"
        Me.txtRegExp.Size = New System.Drawing.Size(470, 20)
        Me.txtRegExp.TabIndex = 10
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(11, 271)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(71, 13)
        Me.Label8.TabIndex = 11
        Me.Label8.Text = "Tipo consulta"
        '
        'cboTipoConsulta
        '
        Me.cboTipoConsulta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTipoConsulta.FormattingEnabled = True
        Me.cboTipoConsulta.Items.AddRange(New Object() {"Empieza por", "Incluye"})
        Me.cboTipoConsulta.Location = New System.Drawing.Point(105, 268)
        Me.cboTipoConsulta.Name = "cboTipoConsulta"
        Me.cboTipoConsulta.Size = New System.Drawing.Size(121, 21)
        Me.cboTipoConsulta.TabIndex = 12
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(12, 9)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(65, 13)
        Me.Label9.TabIndex = 13
        Me.Label9.Text = "GENERAL"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(12, 33)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(77, 13)
        Me.Label10.TabIndex = 14
        Me.Label10.Text = "Conectando a:"
        '
        'cboConectando
        '
        Me.cboConectando.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboConectando.FormattingEnabled = True
        Me.cboConectando.Items.AddRange(New Object() {"Local", "Remota"})
        Me.cboConectando.Location = New System.Drawing.Point(105, 30)
        Me.cboConectando.Name = "cboConectando"
        Me.cboConectando.Size = New System.Drawing.Size(121, 21)
        Me.cboConectando.TabIndex = 15
        '
        'cboMuestraMicro
        '
        Me.cboMuestraMicro.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboMuestraMicro.FormattingEnabled = True
        Me.cboMuestraMicro.Items.AddRange(New Object() {"Vinculados", "NO vinculados"})
        Me.cboMuestraMicro.Location = New System.Drawing.Point(105, 295)
        Me.cboMuestraMicro.Name = "cboMuestraMicro"
        Me.cboMuestraMicro.Size = New System.Drawing.Size(121, 21)
        Me.cboMuestraMicro.TabIndex = 16
        '
        'lblMuestraMicro
        '
        Me.lblMuestraMicro.AutoSize = True
        Me.lblMuestraMicro.Location = New System.Drawing.Point(12, 298)
        Me.lblMuestraMicro.Name = "lblMuestraMicro"
        Me.lblMuestraMicro.Size = New System.Drawing.Size(80, 13)
        Me.lblMuestraMicro.TabIndex = 17
        Me.lblMuestraMicro.Text = "Muestra - Micro"
        '
        'btnProbarConexionLocal
        '
        Me.btnProbarConexionLocal.Location = New System.Drawing.Point(253, 84)
        Me.btnProbarConexionLocal.Name = "btnProbarConexionLocal"
        Me.btnProbarConexionLocal.Size = New System.Drawing.Size(97, 23)
        Me.btnProbarConexionLocal.TabIndex = 18
        Me.btnProbarConexionLocal.Text = "Probar Conexión"
        Me.btnProbarConexionLocal.UseVisualStyleBackColor = True
        '
        'btnProbarConexionRemota
        '
        Me.btnProbarConexionRemota.Location = New System.Drawing.Point(253, 173)
        Me.btnProbarConexionRemota.Name = "btnProbarConexionRemota"
        Me.btnProbarConexionRemota.Size = New System.Drawing.Size(97, 23)
        Me.btnProbarConexionRemota.TabIndex = 19
        Me.btnProbarConexionRemota.Text = "Probar Conexión"
        Me.btnProbarConexionRemota.UseVisualStyleBackColor = True
        '
        'DialogoConfiguracion
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(590, 325)
        Me.Controls.Add(Me.btnProbarConexionRemota)
        Me.Controls.Add(Me.btnProbarConexionLocal)
        Me.Controls.Add(Me.lblMuestraMicro)
        Me.Controls.Add(Me.cboMuestraMicro)
        Me.Controls.Add(Me.cboConectando)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.cboTipoConsulta)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.txtRegExp)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtDSN)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtcnLocal)
        Me.Controls.Add(Me.cboTipoLocal)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialogoConfiguracion"
        Me.Text = "DialogoConfiguracion"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboTipoLocal As System.Windows.Forms.ComboBox
    Friend WithEvents txtcnLocal As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtDSN As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtRegExp As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cboTipoConsulta As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cboConectando As System.Windows.Forms.ComboBox
    Friend WithEvents cboMuestraMicro As System.Windows.Forms.ComboBox
    Friend WithEvents lblMuestraMicro As System.Windows.Forms.Label
    Friend WithEvents btnProbarConexionLocal As System.Windows.Forms.Button
    Friend WithEvents btnProbarConexionRemota As System.Windows.Forms.Button
End Class
