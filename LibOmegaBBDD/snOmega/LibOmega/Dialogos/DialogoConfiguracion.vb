Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.Odbc

Public Class DialogoConfiguracion

    Dim mstrNombreArchivoConfig As String
    Dim mobjConfig As clsConfig

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        mstrNombreArchivoConfig = clsUtil.DLLPath(True) + "Config.xml"

        Try
            If My.Computer.FileSystem.FileExists(mstrNombreArchivoConfig) Then CargarConfiguracion()
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al cargar la configuración", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    ' ************************************************************************
    ' CargarConfiguracion
    ' Desc: Rutina que carga la configuración de BBDD
    ' NBL: 11/01/2007
    ' ************************************************************************
    Private Sub CargarConfiguracion()

        Dim xmlReader As New XmlTextReader(mstrNombreArchivoConfig)
        Dim Reader As New Serialization.XmlSerializer(GetType(clsConfig))

        mobjConfig = Reader.Deserialize(xmlReader)

        ' Cargamos los datos
        Me.cboConectando.SelectedIndex = mobjConfig.Conexion
        Me.cboTipoLocal.SelectedIndex = mobjConfig.TipoLocal
        Me.txtcnLocal.Text = mobjConfig.cnLocal
        Me.txtDSN.Text = mobjConfig.DSNRemota
        Me.txtRegExp.Text = mobjConfig.reHC
        Me.cboTipoConsulta.SelectedIndex = mobjConfig.TipoConsulta
        Me.cboMuestraMicro.SelectedIndex = mobjConfig.MuestraMicro

        xmlReader.Close()

    End Sub

    Private Sub DialogoConfiguracion_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If MessageBox.Show("¿Guardar la configuración?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
            Try
                GuardarConfiguracion()
            Catch ex As Exception
                MessageBox.Show("Ha habido algún error al guardar la configuración" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, _
                                                 MessageBoxIcon.Exclamation)
            End Try
        End If

        My.Settings.DialogoConfiguracionLocation = Me.Location
        My.Settings.Save()

    End Sub

    ' ************************************************************************
    ' GuardarConfiguracion
    ' Desc: Rutina que guarda la configuración de BBDD
    ' NBL: 16/01/2007
    ' ************************************************************************
    Private Sub GuardarConfiguracion()

        Try
            ' En primer lugar miramos si existe el archivo de configuración y lo borramos.
            If My.Computer.FileSystem.FileExists(mstrNombreArchivoConfig) Then _
                My.Computer.FileSystem.DeleteFile(mstrNombreArchivoConfig)
        Catch ex As Exception
            ' Aquí no hago nada porque aunque no lo borre puede que lo pueda guardar igualmente
        End Try

        If mobjConfig Is Nothing Then mobjConfig = New clsConfig

        mobjConfig.Conexion = Me.cboConectando.SelectedIndex
        mobjConfig.TipoLocal = Me.cboTipoLocal.SelectedIndex
        mobjConfig.cnLocal = Me.txtcnLocal.Text.Trim
        mobjConfig.DSNRemota = Me.txtDSN.Text.Trim
        mobjConfig.reHC = Me.txtRegExp.Text.Trim
        mobjConfig.TipoConsulta = Me.cboTipoConsulta.SelectedIndex
        mobjConfig.MuestraMicro = Me.cboMuestraMicro.SelectedIndex

        'Crear un objeto serializado para la clase contactos
        Dim objWriter As New Serialization.XmlSerializer(GetType(clsConfig))
        'Crear un objeto file de tipo StremWriter para almacenar el documento xml
        Dim objFile As New StreamWriter(mstrNombreArchivoConfig)
        'Serializar y crear el documento XML
        objWriter.Serialize(objFile, mobjConfig)
        'Cerrar el archivo
        objFile.Close()

    End Sub

    Private Sub DialogoConfiguracion_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Location = My.Settings.DialogoConfiguracionLocation

    End Sub

    Private Sub cboTipoConsulta_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTipoConsulta.SelectedIndexChanged

    End Sub

    Private Sub btnProbarConexionLocal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProbarConexionLocal.Click

        ' *******************************************************
        ' Probamos si la conexión local está bien configurada
        ' *******************************************************
        If Me.cboTipoLocal.SelectedIndex = -1 Or Me.txtcnLocal.Text.Length = 0 Then
            MessageBox.Show("Configure la conexión local", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        If Me.cboTipoLocal.SelectedIndex = 0 Then
            ' Conexión a Access
            Dim mobjCnnAccess As New OleDbConnection(Me.txtcnLocal.Text)

            Try
                mobjCnnAccess.Open()
                MessageBox.Show("Se ha conectado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                mobjCnnAccess.Close()
            Catch ex As Exception
                MessageBox.Show("Conexión fallida" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        Else
            ' Conexión SQL Server
            Dim mobjCnnSqlServer As New SqlConnection(Me.txtcnLocal.Text)

            Try
                mobjCnnSqlServer.Open()
                MessageBox.Show("Se ha conectado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                mobjCnnSqlServer.Close()
            Catch ex As Exception
                MessageBox.Show("Conexión fallida" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End If


    End Sub

    Private Sub btnProbarConexionRemota_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProbarConexionRemota.Click

        If Me.txtDSN.Text.Length = 0 Then
            MessageBox.Show("Indique un nombre DSN", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        Try
            Dim mobjCnnRemota As New OdbcConnection("DSN=" + Me.txtDSN.Text)
            mobjCnnRemota.Open()
            MessageBox.Show("Se ha conectado correctamente", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            mobjCnnRemota.Close()
        Catch ex As Exception
            MessageBox.Show("Conexión fallida" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

End Class