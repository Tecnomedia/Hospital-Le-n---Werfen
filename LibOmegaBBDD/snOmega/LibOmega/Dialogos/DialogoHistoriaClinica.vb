Imports System.Windows.Forms

Public Class DialogoHistoriaClinica

    Private mobjConsulta As clsConsulta
    Public Resultado As String

    ' Instancia a la clase de reordenación de columnas de un listview
    Private lvwColumnSorter As ListViewColumnSorter

    Sub New(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, ByVal pstrTextoBusqueda As String, _
                        Optional ByVal parlstPreResultado As ArrayList = Nothing)

        InitializeComponent()

        Me.Text = pstrTituloDialogo
        Me.lblInfo.Text = pstrLabel

        mobjConsulta = New clsConsulta

        If Not parlstPreResultado Is Nothing Then
            CargarResultado(parlstPreResultado)
        End If

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter = New ListViewColumnSorter()
        Me.lvResultadoHC.ListViewItemSorter = lvwColumnSorter

        ' Mostramos los datos ordenados de los apellidos
        lvwColumnSorter.SortColumn = 1
        lvwColumnSorter.Order = SortOrder.Ascending

    End Sub

    ' **************************************************************
    ' CargarResultado
    ' Desc: Rutina que carga el resultado de una busqueda
    ' 27/2/2007
    ' **************************************************************
    Private Sub CargarResultado(ByVal parlstResultado As ArrayList)

        Me.Cursor = Cursors.WaitCursor

        If parlstResultado.Count > 0 Then
            With Me.lvResultadoHC
                .Items.Clear()
                .BeginUpdate()
                .SuspendLayout()
                .Items.AddRange(parlstResultado.ToArray(GetType(ListViewItem)))
                .EndUpdate()
                .ResumeLayout()
                .TabIndex = 0
                .Focus()
                .Items(0).Selected = True
                .Items(0).Focused = True
            End With
        Else
            MessageBox.Show("Ningún resultado", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub txtTextoBusqueda_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTextoBusqueda.KeyDown

        If e.KeyCode = Keys.Enter And Me.txtTextoBusqueda.Text.Trim.Length > 0 Then
            Dim larlstResultado As ArrayList = Me.mobjConsulta.BuscaHC(Me.txtTextoBusqueda.Text.Trim)
            If Not larlstResultado Is Nothing Then
                CargarResultado(larlstResultado)
            End If
        End If

    End Sub

    Private Sub txtTextoBusqueda_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTextoBusqueda.TextChanged

    End Sub

    Private Sub DialogoServicios_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Me.Top > 0 Then
            My.Settings.DialogoHCLocation = Me.Location
        End If

        If Me.WindowState <> FormWindowState.Maximized Then My.Settings.DialogoHCSize = Me.Size
        My.Settings.DialogoHCState = Me.WindowState

        My.Settings.DialogoHCNumeroWidth = Me.lvResultadoHC.Columns(0).Width
        My.Settings.DialogoHCApellidosWidth = Me.lvResultadoHC.Columns(1).Width
        My.Settings.DialogoHCNombreWidth = Me.lvResultadoHC.Columns(2).Width
        My.Settings.DialogoHCFechaNacimientoWidth = Me.lvResultadoHC.Columns(3).Width

        My.Settings.Save()

    End Sub

    Private Sub DialogoServicios_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.F12 Then
            AceptarResultado()
        End If

        If e.KeyCode = Keys.F12 Or e.KeyCode = Keys.Escape Then Me.Close()

    End Sub

    Private Sub AceptarResultado()

        If Me.lvResultadoHC.SelectedItems.Count > 0 Then
            With Me.lvResultadoHC.SelectedItems(0)
                Me.Resultado = .Tag
            End With
            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If

    End Sub

    Private Sub DialogoServicios_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not My.Settings.DialogoHCLocation.IsEmpty Then
            Me.Location = My.Settings.DialogoHCLocation
            If Me.Top < 0 Then
                Me.Top = 0
                Me.Left = 0
            End If
        End If

        Me.Size = My.Settings.DialogoHCSize

        If My.Settings.DialogoHCState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If

        If Me.Width < 300 Then Me.Width = 300
        If Me.Height < 300 Then Me.Height = 300

        Me.lvResultadoHC.Columns(0).Width = IIf(My.Settings.DialogoHCNumeroWidth = 0, 50, My.Settings.DialogoHCNumeroWidth)
        Me.lvResultadoHC.Columns(1).Width = IIf(My.Settings.DialogoHCApellidosWidth = 0, 50, My.Settings.DialogoHCApellidosWidth)
        Me.lvResultadoHC.Columns(2).Width = IIf(My.Settings.DialogoHCNombreWidth = 0, 50, My.Settings.DialogoHCNombreWidth)
        Me.lvResultadoHC.Columns(3).Width = IIf(My.Settings.DialogoHCFechaNacimientoWidth = 0, 50, My.Settings.DialogoHCFechaNacimientoWidth)

    End Sub

    Private Sub lvwResultadoMedicos_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvResultadoHC.ColumnClick

        ' Indicamos que no es una fecha
        If e.Column = 3 Then
            lvwColumnSorter.OrderDate = True
        Else
            lvwColumnSorter.OrderDate = False
        End If

        ' Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.lvResultadoHC.Sort()
    End Sub

    Private Sub lvResultadoHC_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvResultadoHC.DoubleClick

        AceptarResultado()

    End Sub

    Private Sub lvwResultadoServicios_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lvResultadoHC.SelectedIndexChanged

    End Sub

End Class