<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogoDestinos
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
        Me.lvResultadoDestinos = New System.Windows.Forms.ListView
        Me.Codigo = New System.Windows.Forms.ColumnHeader
        Me.Nombre = New System.Windows.Forms.ColumnHeader
        Me.txtTextoBusqueda = New System.Windows.Forms.TextBox
        Me.lblInfo = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'lvResultadoDestinos
        '
        Me.lvResultadoDestinos.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvResultadoDestinos.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Codigo, Me.Nombre})
        Me.lvResultadoDestinos.FullRowSelect = True
        Me.lvResultadoDestinos.GridLines = True
        Me.lvResultadoDestinos.HideSelection = False
        Me.lvResultadoDestinos.Location = New System.Drawing.Point(12, 51)
        Me.lvResultadoDestinos.MultiSelect = False
        Me.lvResultadoDestinos.Name = "lvResultadoDestinos"
        Me.lvResultadoDestinos.Size = New System.Drawing.Size(336, 287)
        Me.lvResultadoDestinos.TabIndex = 0
        Me.lvResultadoDestinos.UseCompatibleStateImageBehavior = False
        Me.lvResultadoDestinos.View = System.Windows.Forms.View.Details
        '
        'Codigo
        '
        Me.Codigo.Text = "Código"
        '
        'Nombre
        '
        Me.Nombre.Text = "Nombre"
        '
        'txtTextoBusqueda
        '
        Me.txtTextoBusqueda.Location = New System.Drawing.Point(12, 27)
        Me.txtTextoBusqueda.Name = "txtTextoBusqueda"
        Me.txtTextoBusqueda.Size = New System.Drawing.Size(174, 20)
        Me.txtTextoBusqueda.TabIndex = 1
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
        'DialogoDestinos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(360, 350)
        Me.Controls.Add(Me.lblInfo)
        Me.Controls.Add(Me.txtTextoBusqueda)
        Me.Controls.Add(Me.lvResultadoDestinos)
        Me.KeyPreview = True
        Me.Name = "DialogoDestinos"
        Me.Text = "DialogoDestinos"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvResultadoDestinos As System.Windows.Forms.ListView
    Friend WithEvents txtTextoBusqueda As System.Windows.Forms.TextBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents Codigo As System.Windows.Forms.ColumnHeader
    Friend WithEvents Nombre As System.Windows.Forms.ColumnHeader
End Class
