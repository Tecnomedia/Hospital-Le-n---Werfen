<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogoHistoriaClinica
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
        Me.lvResultadoHC = New System.Windows.Forms.ListView
        Me.NumHC = New System.Windows.Forms.ColumnHeader
        Me.Apellidos = New System.Windows.Forms.ColumnHeader
        Me.Nombre = New System.Windows.Forms.ColumnHeader
        Me.FechaNac = New System.Windows.Forms.ColumnHeader
        Me.txtTextoBusqueda = New System.Windows.Forms.TextBox
        Me.lblInfo = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lvResultadoHC
        '
        Me.lvResultadoHC.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvResultadoHC.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.NumHC, Me.Apellidos, Me.Nombre, Me.FechaNac})
        Me.lvResultadoHC.FullRowSelect = True
        Me.lvResultadoHC.GridLines = True
        Me.lvResultadoHC.HideSelection = False
        Me.lvResultadoHC.Location = New System.Drawing.Point(12, 51)
        Me.lvResultadoHC.MultiSelect = False
        Me.lvResultadoHC.Name = "lvResultadoHC"
        Me.lvResultadoHC.Size = New System.Drawing.Size(336, 287)
        Me.lvResultadoHC.TabIndex = 1
        Me.lvResultadoHC.UseCompatibleStateImageBehavior = False
        Me.lvResultadoHC.View = System.Windows.Forms.View.Details
        '
        'NumHC
        '
        Me.NumHC.Text = "HC"
        '
        'Apellidos
        '
        Me.Apellidos.Text = "Apellidos"
        '
        'Nombre
        '
        Me.Nombre.Text = "Nombre"
        '
        'FechaNac
        '
        Me.FechaNac.Text = "FechaNacimiento"
        '
        'txtTextoBusqueda
        '
        Me.txtTextoBusqueda.Location = New System.Drawing.Point(12, 27)
        Me.txtTextoBusqueda.Name = "txtTextoBusqueda"
        Me.txtTextoBusqueda.Size = New System.Drawing.Size(174, 20)
        Me.txtTextoBusqueda.TabIndex = 0
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.Location = New System.Drawing.Point(12, 9)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(39, 13)
        Me.lblInfo.TabIndex = 2
        Me.lblInfo.Text = "Label1"
        '
        'DialogoHistoriaClinica
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(360, 350)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.txtTextoBusqueda)
        Me.Controls.Add(Me.lvResultadoHC)
        Me.KeyPreview = True
        Me.Name = "DialogoHistoriaClinica"
        Me.Text = "DialogoHistoriasClinicas"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvResultadoHC As System.Windows.Forms.ListView
    Friend WithEvents txtTextoBusqueda As System.Windows.Forms.TextBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents NumHC As System.Windows.Forms.ColumnHeader
    Friend WithEvents Apellidos As System.Windows.Forms.ColumnHeader
    Friend WithEvents Nombre As System.Windows.Forms.ColumnHeader
    Friend WithEvents FechaNac As System.Windows.Forms.ColumnHeader
End Class
