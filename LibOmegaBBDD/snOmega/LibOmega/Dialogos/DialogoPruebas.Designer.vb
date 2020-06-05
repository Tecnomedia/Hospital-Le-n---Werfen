<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogoPruebas
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DialogoPruebas))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.lblSeleccionadas = New System.Windows.Forms.Label
        Me.btnCancelar = New System.Windows.Forms.Button
        Me.btnAceptar = New System.Windows.Forms.Button
        Me.txtBuscaMicro = New System.Windows.Forms.TextBox
        Me.txtNombreMuestra = New System.Windows.Forms.TextBox
        Me.txtAbrvMuestra = New System.Windows.Forms.TextBox
        Me.txtCodigoMuestra = New System.Windows.Forms.TextBox
        Me.txtBuscaMuestraMicro = New System.Windows.Forms.TextBox
        Me.lblMicrobiologia = New System.Windows.Forms.Label
        Me.txtBuscaBioquimica = New System.Windows.Forms.TextBox
        Me.lblBioquimica = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.lvwResultadoBioquimica = New System.Windows.Forms.ListView
        Me.Codigo = New System.Windows.Forms.ColumnHeader
        Me.Abrv = New System.Windows.Forms.ColumnHeader
        Me.Nombre = New System.Windows.Forms.ColumnHeader
        Me.Bioquimica = New System.Windows.Forms.ImageList(Me.components)
        Me.ilPerfilPrueba = New System.Windows.Forms.ImageList(Me.components)
        Me.lvwResultadoMuestras = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.lvwResultadoMicro = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.Microbiologia = New System.Windows.Forms.ImageList(Me.components)
        Me.dgvPruebasTotal = New System.Windows.Forms.DataGridView
        Me.CodigoT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AbrvT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NombreT = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.CodigoM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.AbrvM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.NombreM = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.TipoPrueba = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.PruebaPerfil = New System.Windows.Forms.DataGridViewTextBoxColumn
        Me.btnBorrarTodo = New System.Windows.Forms.Button
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvPruebasTotal, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblSeleccionadas
        '
        Me.lblSeleccionadas.AutoSize = True
        Me.lblSeleccionadas.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSeleccionadas.Location = New System.Drawing.Point(9, 114)
        Me.lblSeleccionadas.Name = "lblSeleccionadas"
        Me.lblSeleccionadas.Size = New System.Drawing.Size(77, 16)
        Me.lblSeleccionadas.TabIndex = 2
        Me.lblSeleccionadas.Text = "Selección"
        '
        'btnCancelar
        '
        Me.btnCancelar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCancelar.Location = New System.Drawing.Point(602, 539)
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(60, 20)
        Me.btnCancelar.TabIndex = 9
        Me.btnCancelar.Text = "&Cancelar"
        Me.btnCancelar.UseVisualStyleBackColor = True
        '
        'btnAceptar
        '
        Me.btnAceptar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAceptar.Location = New System.Drawing.Point(532, 539)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(64, 20)
        Me.btnAceptar.TabIndex = 10
        Me.btnAceptar.Text = "&Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'txtBuscaMicro
        '
        Me.txtBuscaMicro.Location = New System.Drawing.Point(70, 79)
        Me.txtBuscaMicro.Name = "txtBuscaMicro"
        Me.txtBuscaMicro.Size = New System.Drawing.Size(121, 21)
        Me.txtBuscaMicro.TabIndex = 2
        '
        'txtNombreMuestra
        '
        Me.txtNombreMuestra.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtNombreMuestra.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtNombreMuestra.Location = New System.Drawing.Point(292, 53)
        Me.txtNombreMuestra.Name = "txtNombreMuestra"
        Me.txtNombreMuestra.ReadOnly = True
        Me.txtNombreMuestra.Size = New System.Drawing.Size(370, 21)
        Me.txtNombreMuestra.TabIndex = 30
        '
        'txtAbrvMuestra
        '
        Me.txtAbrvMuestra.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtAbrvMuestra.Location = New System.Drawing.Point(236, 53)
        Me.txtAbrvMuestra.Name = "txtAbrvMuestra"
        Me.txtAbrvMuestra.ReadOnly = True
        Me.txtAbrvMuestra.Size = New System.Drawing.Size(50, 21)
        Me.txtAbrvMuestra.TabIndex = 29
        '
        'txtCodigoMuestra
        '
        Me.txtCodigoMuestra.BackColor = System.Drawing.Color.LemonChiffon
        Me.txtCodigoMuestra.Location = New System.Drawing.Point(197, 53)
        Me.txtCodigoMuestra.Name = "txtCodigoMuestra"
        Me.txtCodigoMuestra.ReadOnly = True
        Me.txtCodigoMuestra.Size = New System.Drawing.Size(33, 21)
        Me.txtCodigoMuestra.TabIndex = 28
        '
        'txtBuscaMuestraMicro
        '
        Me.txtBuscaMuestraMicro.Location = New System.Drawing.Point(70, 53)
        Me.txtBuscaMuestraMicro.Name = "txtBuscaMuestraMicro"
        Me.txtBuscaMuestraMicro.Size = New System.Drawing.Size(121, 21)
        Me.txtBuscaMuestraMicro.TabIndex = 1
        '
        'lblMicrobiologia
        '
        Me.lblMicrobiologia.AutoSize = True
        Me.lblMicrobiologia.BackColor = System.Drawing.Color.LemonChiffon
        Me.lblMicrobiologia.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMicrobiologia.Location = New System.Drawing.Point(9, 54)
        Me.lblMicrobiologia.Name = "lblMicrobiologia"
        Me.lblMicrobiologia.Size = New System.Drawing.Size(56, 16)
        Me.lblMicrobiologia.TabIndex = 23
        Me.lblMicrobiologia.Text = "MICRO"
        '
        'txtBuscaBioquimica
        '
        Me.txtBuscaBioquimica.BackColor = System.Drawing.Color.White
        Me.txtBuscaBioquimica.Location = New System.Drawing.Point(70, 12)
        Me.txtBuscaBioquimica.Name = "txtBuscaBioquimica"
        Me.txtBuscaBioquimica.Size = New System.Drawing.Size(120, 21)
        Me.txtBuscaBioquimica.TabIndex = 0
        '
        'lblBioquimica
        '
        Me.lblBioquimica.AutoSize = True
        Me.lblBioquimica.BackColor = System.Drawing.Color.Moccasin
        Me.lblBioquimica.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBioquimica.Location = New System.Drawing.Point(10, 14)
        Me.lblBioquimica.Name = "lblBioquimica"
        Me.lblBioquimica.Size = New System.Drawing.Size(44, 16)
        Me.lblBioquimica.TabIndex = 38
        Me.lblBioquimica.Text = "BIOQ"
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackColor = System.Drawing.Color.Moccasin
        Me.PictureBox1.Location = New System.Drawing.Point(-3, -12)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(675, 54)
        Me.PictureBox1.TabIndex = 46
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.BackColor = System.Drawing.Color.LemonChiffon
        Me.PictureBox2.Location = New System.Drawing.Point(-3, 37)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(675, 74)
        Me.PictureBox2.TabIndex = 47
        Me.PictureBox2.TabStop = False
        '
        'lvwResultadoBioquimica
        '
        Me.lvwResultadoBioquimica.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwResultadoBioquimica.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Codigo, Me.Abrv, Me.Nombre})
        Me.lvwResultadoBioquimica.FullRowSelect = True
        Me.lvwResultadoBioquimica.GridLines = True
        Me.lvwResultadoBioquimica.HideSelection = False
        Me.lvwResultadoBioquimica.Location = New System.Drawing.Point(8, 37)
        Me.lvwResultadoBioquimica.MultiSelect = False
        Me.lvwResultadoBioquimica.Name = "lvwResultadoBioquimica"
        Me.lvwResultadoBioquimica.Size = New System.Drawing.Size(652, 520)
        Me.lvwResultadoBioquimica.SmallImageList = Me.Bioquimica
        Me.lvwResultadoBioquimica.TabIndex = 48
        Me.lvwResultadoBioquimica.UseCompatibleStateImageBehavior = False
        Me.lvwResultadoBioquimica.View = System.Windows.Forms.View.Details
        Me.lvwResultadoBioquimica.Visible = False
        '
        'Codigo
        '
        Me.Codigo.Text = "Código"
        '
        'Abrv
        '
        Me.Abrv.Text = "Abrv"
        '
        'Nombre
        '
        Me.Nombre.Text = "Nombre"
        '
        'Bioquimica
        '
        Me.Bioquimica.ImageStream = CType(resources.GetObject("Bioquimica.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.Bioquimica.TransparentColor = System.Drawing.Color.Transparent
        Me.Bioquimica.Images.SetKeyName(0, "BIO_PRUEBA.ICO")
        Me.Bioquimica.Images.SetKeyName(1, "BIO_PERFIL.ICO")
        '
        'ilPerfilPrueba
        '
        Me.ilPerfilPrueba.ImageStream = CType(resources.GetObject("ilPerfilPrueba.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ilPerfilPrueba.TransparentColor = System.Drawing.Color.Transparent
        Me.ilPerfilPrueba.Images.SetKeyName(0, "PRU.ICO")
        Me.ilPerfilPrueba.Images.SetKeyName(1, "PER.ICO")
        Me.ilPerfilPrueba.Images.SetKeyName(2, "MUE.ICO")
        '
        'lvwResultadoMuestras
        '
        Me.lvwResultadoMuestras.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwResultadoMuestras.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3})
        Me.lvwResultadoMuestras.FullRowSelect = True
        Me.lvwResultadoMuestras.GridLines = True
        Me.lvwResultadoMuestras.HideSelection = False
        Me.lvwResultadoMuestras.Location = New System.Drawing.Point(8, 79)
        Me.lvwResultadoMuestras.MultiSelect = False
        Me.lvwResultadoMuestras.Name = "lvwResultadoMuestras"
        Me.lvwResultadoMuestras.Size = New System.Drawing.Size(652, 479)
        Me.lvwResultadoMuestras.SmallImageList = Me.ilPerfilPrueba
        Me.lvwResultadoMuestras.TabIndex = 51
        Me.lvwResultadoMuestras.UseCompatibleStateImageBehavior = False
        Me.lvwResultadoMuestras.View = System.Windows.Forms.View.Details
        Me.lvwResultadoMuestras.Visible = False
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Código"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Abrv"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Nombre"
        '
        'lvwResultadoMicro
        '
        Me.lvwResultadoMicro.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvwResultadoMicro.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6})
        Me.lvwResultadoMicro.FullRowSelect = True
        Me.lvwResultadoMicro.GridLines = True
        Me.lvwResultadoMicro.HideSelection = False
        Me.lvwResultadoMicro.Location = New System.Drawing.Point(8, 105)
        Me.lvwResultadoMicro.MultiSelect = False
        Me.lvwResultadoMicro.Name = "lvwResultadoMicro"
        Me.lvwResultadoMicro.Size = New System.Drawing.Size(652, 452)
        Me.lvwResultadoMicro.SmallImageList = Me.Microbiologia
        Me.lvwResultadoMicro.TabIndex = 52
        Me.lvwResultadoMicro.UseCompatibleStateImageBehavior = False
        Me.lvwResultadoMicro.View = System.Windows.Forms.View.Details
        Me.lvwResultadoMicro.Visible = False
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Código"
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Abrv"
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Nombre"
        '
        'Microbiologia
        '
        Me.Microbiologia.ImageStream = CType(resources.GetObject("Microbiologia.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.Microbiologia.TransparentColor = System.Drawing.Color.Transparent
        Me.Microbiologia.Images.SetKeyName(0, "MIC_PRUEBA.ICO")
        Me.Microbiologia.Images.SetKeyName(1, "MIC_PERFIL.ICO")
        '
        'dgvPruebasTotal
        '
        Me.dgvPruebasTotal.AllowUserToAddRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.Transparent
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        Me.dgvPruebasTotal.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvPruebasTotal.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvPruebasTotal.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.dgvPruebasTotal.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvPruebasTotal.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.CodigoT, Me.AbrvT, Me.NombreT, Me.CodigoM, Me.AbrvM, Me.NombreM, Me.TipoPrueba, Me.PruebaPerfil})
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Transparent
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.Transparent
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvPruebasTotal.DefaultCellStyle = DataGridViewCellStyle4
        Me.dgvPruebasTotal.GridColor = System.Drawing.SystemColors.MenuBar
        Me.dgvPruebasTotal.Location = New System.Drawing.Point(8, 133)
        Me.dgvPruebasTotal.MultiSelect = False
        Me.dgvPruebasTotal.Name = "dgvPruebasTotal"
        Me.dgvPruebasTotal.ReadOnly = True
        Me.dgvPruebasTotal.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvPruebasTotal.Size = New System.Drawing.Size(654, 400)
        Me.dgvPruebasTotal.TabIndex = 53
        '
        'CodigoT
        '
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Transparent
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Transparent
        Me.CodigoT.DefaultCellStyle = DataGridViewCellStyle2
        Me.CodigoT.HeaderText = "Código"
        Me.CodigoT.Name = "CodigoT"
        Me.CodigoT.ReadOnly = True
        '
        'AbrvT
        '
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AbrvT.DefaultCellStyle = DataGridViewCellStyle3
        Me.AbrvT.HeaderText = "Abrv"
        Me.AbrvT.Name = "AbrvT"
        Me.AbrvT.ReadOnly = True
        '
        'NombreT
        '
        Me.NombreT.HeaderText = "Nombre"
        Me.NombreT.Name = "NombreT"
        Me.NombreT.ReadOnly = True
        '
        'CodigoM
        '
        Me.CodigoM.HeaderText = "C. Muestra"
        Me.CodigoM.Name = "CodigoM"
        Me.CodigoM.ReadOnly = True
        '
        'AbrvM
        '
        Me.AbrvM.HeaderText = "A. Muestra"
        Me.AbrvM.Name = "AbrvM"
        Me.AbrvM.ReadOnly = True
        '
        'NombreM
        '
        Me.NombreM.HeaderText = "N. Muestra"
        Me.NombreM.Name = "NombreM"
        Me.NombreM.ReadOnly = True
        '
        'TipoPrueba
        '
        Me.TipoPrueba.HeaderText = "TipoPrueba"
        Me.TipoPrueba.Name = "TipoPrueba"
        Me.TipoPrueba.ReadOnly = True
        Me.TipoPrueba.Visible = False
        '
        'PruebaPerfil
        '
        Me.PruebaPerfil.HeaderText = "PruebaPerfil"
        Me.PruebaPerfil.Name = "PruebaPerfil"
        Me.PruebaPerfil.ReadOnly = True
        Me.PruebaPerfil.Visible = False
        '
        'btnBorrarTodo
        '
        Me.btnBorrarTodo.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnBorrarTodo.Location = New System.Drawing.Point(8, 539)
        Me.btnBorrarTodo.Name = "btnBorrarTodo"
        Me.btnBorrarTodo.Size = New System.Drawing.Size(71, 20)
        Me.btnBorrarTodo.TabIndex = 54
        Me.btnBorrarTodo.Text = "&Borrar todo"
        Me.btnBorrarTodo.UseVisualStyleBackColor = True
        '
        'DialogoPruebas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.LightGray
        Me.ClientSize = New System.Drawing.Size(670, 565)
        Me.Controls.Add(Me.btnBorrarTodo)
        Me.Controls.Add(Me.dgvPruebasTotal)
        Me.Controls.Add(Me.txtBuscaBioquimica)
        Me.Controls.Add(Me.lblBioquimica)
        Me.Controls.Add(Me.txtBuscaMicro)
        Me.Controls.Add(Me.txtNombreMuestra)
        Me.Controls.Add(Me.txtAbrvMuestra)
        Me.Controls.Add(Me.txtCodigoMuestra)
        Me.Controls.Add(Me.txtBuscaMuestraMicro)
        Me.Controls.Add(Me.lblMicrobiologia)
        Me.Controls.Add(Me.btnCancelar)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.lblSeleccionadas)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.lvwResultadoMuestras)
        Me.Controls.Add(Me.lvwResultadoMicro)
        Me.Controls.Add(Me.lvwResultadoBioquimica)
        Me.KeyPreview = True
        Me.Name = "DialogoPruebas"
        Me.Text = "DialogoPruebas"
        Me.TopMost = True
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvPruebasTotal, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblSeleccionadas As System.Windows.Forms.Label
    Friend WithEvents btnCancelar As System.Windows.Forms.Button
    Friend WithEvents btnAceptar As System.Windows.Forms.Button
    Friend WithEvents txtBuscaMicro As System.Windows.Forms.TextBox
    Friend WithEvents txtNombreMuestra As System.Windows.Forms.TextBox
    Friend WithEvents txtAbrvMuestra As System.Windows.Forms.TextBox
    Friend WithEvents txtCodigoMuestra As System.Windows.Forms.TextBox
    Friend WithEvents txtBuscaMuestraMicro As System.Windows.Forms.TextBox
    Friend WithEvents lblMicrobiologia As System.Windows.Forms.Label
    Friend WithEvents txtBuscaBioquimica As System.Windows.Forms.TextBox
    Friend WithEvents lblBioquimica As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents lvwResultadoBioquimica As System.Windows.Forms.ListView
    Friend WithEvents Codigo As System.Windows.Forms.ColumnHeader
    Friend WithEvents Abrv As System.Windows.Forms.ColumnHeader
    Friend WithEvents Nombre As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvwResultadoMuestras As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lvwResultadoMicro As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents dgvPruebasTotal As System.Windows.Forms.DataGridView
    Friend WithEvents ilPerfilPrueba As System.Windows.Forms.ImageList
    Friend WithEvents CodigoT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AbrvT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NombreT As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CodigoM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents AbrvM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NombreM As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents TipoPrueba As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PruebaPerfil As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btnBorrarTodo As System.Windows.Forms.Button
    Friend WithEvents Bioquimica As System.Windows.Forms.ImageList
    Friend WithEvents Microbiologia As System.Windows.Forms.ImageList
End Class
