<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmProgressExport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.lstLogExportacion = New System.Windows.Forms.ListBox
        Me.prgEstadoProceso = New System.Windows.Forms.ProgressBar
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lstLogExportacion
        '
        Me.lstLogExportacion.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstLogExportacion.FormattingEnabled = True
        Me.lstLogExportacion.Location = New System.Drawing.Point(12, 12)
        Me.lstLogExportacion.Name = "lstLogExportacion"
        Me.lstLogExportacion.ScrollAlwaysVisible = True
        Me.lstLogExportacion.Size = New System.Drawing.Size(654, 251)
        Me.lstLogExportacion.TabIndex = 0
        '
        'prgEstadoProceso
        '
        Me.prgEstadoProceso.Location = New System.Drawing.Point(12, 272)
        Me.prgEstadoProceso.Name = "prgEstadoProceso"
        Me.prgEstadoProceso.Size = New System.Drawing.Size(654, 23)
        Me.prgEstadoProceso.TabIndex = 1
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 500
        '
        'frmProgressExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(678, 305)
        Me.Controls.Add(Me.prgEstadoProceso)
        Me.Controls.Add(Me.lstLogExportacion)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmProgressExport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Progreso exportación ASTM"
        Me.TopMost = True
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lstLogExportacion As System.Windows.Forms.ListBox
    Friend WithEvents prgEstadoProceso As System.Windows.Forms.ProgressBar
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
End Class
