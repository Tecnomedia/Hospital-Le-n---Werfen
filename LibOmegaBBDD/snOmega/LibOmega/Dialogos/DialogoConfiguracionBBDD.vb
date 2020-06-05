Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.Sqlclient
Imports System.Data.Odbc

' Defino los tipos de BBDD que hay, locales o remotas
Public Enum BBDD
    Local
    Remota
End Enum

Friend Class DialogoConfiguracionBBDD

    Dim mstrNombreArchivoConfigBBDD As String
    Dim mobjConfigBBDD As clsConfigBBDD
    Dim TipoBBDD As BBDD
    Dim mobjConfig As clsConfig
    ' Las diferentes conexiones a BBDD
    Dim mobjCnnAccess As OleDbConnection
    Dim mobjCnnSqlServer As SqlConnection
    Dim mobjCnnODBC As OdbcConnection

    Public Sub New(ByVal TipoBBDD As BBDD)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Cargamos la configuración de la aplicación
        Dim mstrNombreArchivoConfig As String = clsUtil.DLLPath(True) + "Config.xml"
        clsUtil.CargarConfiguracion(mobjConfig, mstrNombreArchivoConfig)

        ' Add any initialization after the InitializeComponent() call.
        Me.TipoBBDD = TipoBBDD

        ' Antes de poner la info que hay guardada en la configuración, lo que hacemos es cargar los campos
        ' de las tablas
        Try
            If AbrirConexion() Then
                CargarTablas()
            End If
        Catch ex As Exception
        End Try

        ' Miramos si existe el archivo de configuración, hay dos tipos LocalBBDD.xml y RemotaBBDD
        If TipoBBDD = BBDD.Local Then
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigLocalBBDD.xml"
        Else
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigRemotaBBDD.xml"
        End If

        If My.Computer.FileSystem.FileExists(mstrNombreArchivoConfigBBDD) Then CargarConfiguracion()

        ' Llamamos a la carga de campos de las tablas
        Try
            CargarCampos()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    ' ************************************************************************
    ' ConexionActiva
    ' Desc: Función que comprueba si estamos conectado a la BBDD
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function ConexionActiva() As Boolean

        Try

            If Me.TipoBBDD = BBDD.Local Then
                If Me.mobjConfig.TipoLocal = 0 Then
                    If Me.mobjCnnAccess.State = ConnectionState.Open Then Return True
                Else
                    If Me.mobjCnnSqlServer.State = ConnectionState.Open Then Return True
                End If
            Else
                If Me.mobjCnnODBC.State = ConnectionState.Open Then Return True
            End If

            ' Si llega hasta aquí es que no hay conexión
            Return False

        Catch ex As Exception

            Return False

        End Try

    End Function

    ' ************************************************************************
    ' CargarCampos
    ' Desc: Rutina general que carga todos los combos con nombres de campos 
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Sub CargarCampos()

        If ConexionActiva() Then
            ' Tabla bioquimica
            If Me.cboTablaBioquimica.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaBioquimica.Text, Me.tabBioquimica, _
                                                                                    "cboCodigoBioquimica", "cboAbrvBioquimica", "cboNombreBioquimica")
            If Me.cboTablaMuestra.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaMuestra.Text, Me.tabMuestra, _
                                                                                    "cboCodigoMuestra", "cboAbrvMuestra", "cboNombreMuestra")
            If Me.cboTablaMicro.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaMicro.Text, Me.tabMicrobiologia, _
                                                                                "cboCodigoMicro", "cboAbrvMicro", "cboNombreMicro")
            If Me.cboTablaMicroMuestra.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaMicroMuestra.Text, Me.tabMicroMuestra, _
                                                                                            "cboCodPrueba", "cboCodMuestra")
            If Me.cboTablaPerfilBio.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaPerfilBio.Text, Me.tabPerfilBioquimica, _
                                                                                    "cboCodigoPerfilBio", "cboNombrePerfilBio")
            If Me.cboTablaPerfilMicro.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaPerfilMicro.Text, Me.tabPerfilMicro, _
                                                                                        "cboCodigoPerfilMicro", "cboNombrePerfilMicro", "cboCodMuestraPerfilMicro")
            If Me.cboTablaMedicos.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaMedicos.Text, Me.tabMedicos, _
                                                                                    "cboCodigoMedico", "cboNombreMedico")
            If Me.cboTablaHC.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaHC.Text, Me.tabHC, _
                                                                        "cboIDHC", "cboNumHC", "cboApellidosHC", "cboNombreHC", "cboNumeroSSHC", "cboCodSexoHC", _
                                                                        "cboFechaNacimientoHC", "cboDNIHC", "cboDireccionHC", "cboPoblacionHC", "cboProvinciaHC", "cboCPHC", _
                                                                        "cboTelefonoHC", "cboNombreCompletoHC")
            If Me.cboTablaServicio.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaServicio.Text, Me.tabServicios, _
                                                                                    "cboCodigoServicio", "cboNombreServicio")
            If Me.cboTablaCorrelaciones.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaCorrelaciones.Text, Me.tabCorrelaciones, _
                                                                                        "cboIDCorrelacion", "cboCodigoCampoCorrelacion", "cboCodigoDesencadenanteCorrelacion", _
                                                                                        "cboCodigoDoctorCorrelacion", "cboCodigoServicioCorrelacion", "cboCodigoOrigenCorrelacion", _
                                                                                        "cboCodigoDestinoCorrelacion", "cboCodigoMotivoCorrelacion", "cboCodigoTipoCorrelacion", _
                                                                                        "cboCodigoGrupoFacturacionCorrelacion", "cboCodigoCargoCorrelacion")
            If Me.cboTablaOrigenes.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaOrigenes.Text, Me.tabOrigenes, _
                                                                                    "cboCodigoOrigen", "cboNombreOrigen")
            If Me.cboTablaDestinos.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaDestinos.Text, Me.tabDestinos, _
                                                                                    "cboCodigoDestinos", "cboNombreDestinos")
            If Me.cboTablaDiagnostico.Text.Length > 0 Then CargarNombresCampos(Me.cboTablaDiagnostico.Text, Me.tabDiagnosticos, _
                                                                                                                        "cboCodigoDiagnostico", "cboNombreDiagnostico")
        End If

    End Sub

    ' ************************************************************************
    ' CargarNombresCampos
    ' Desc: Rutina que carga los campos de la tabla pasada como parámetro a los combos del array
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Sub CargarNombresCampos(ByVal pstrNombreTabla As String, ByVal pobjTabPage As TabPage, ByVal ParamArray Combos() As String)

        Dim larrList As ArrayList = ArrayListCampos(pstrNombreTabla)

        For lintContador As Integer = 0 To Combos.Length - 1
            Dim lstrNombreControl As String = Combos(lintContador).ToString()
            Dim mobjCombo As ComboBox = CType(pobjTabPage.Controls.Item(lstrNombreControl), ComboBox)
            mobjCombo.Items.Clear()
            mobjCombo.Items.AddRange(larrList.ToArray())
        Next

    End Sub

    ' ************************************************************************
    ' ArrayListCampos
    ' Desc: Devuelve un arraylist  donde podremos encontrar los campos
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function ArrayListCampos(ByVal pstrNombreTabla As String) As ArrayList

        Dim larrList As New ArrayList
        Dim lobjColumnas As DataColumnCollection

        Try
            If Me.TipoBBDD = BBDD.Local Then
                If Me.mobjConfig.TipoLocal = 0 Then
                    lobjColumnas = DataSetCamposLocalAccess(pstrNombreTabla).Tables(0).Columns
                Else
                    lobjColumnas = DataSetCamposLocalServer(pstrNombreTabla).Tables(0).Columns
                End If
            Else
                lobjColumnas = DataSetCamposRemota(pstrNombreTabla).Tables(0).Columns
            End If
            ' Llenamos el array list
            If lobjColumnas.Count > 0 Then
                For Each lobjColumn As DataColumn In lobjColumnas
                    larrList.Add(lobjColumn.ColumnName)
                Next
            End If
        Catch ex As Exception
            larrList.Clear()
        End Try

        Return larrList

    End Function

    ' ************************************************************************
    ' DataSetCamposLocalAccess
    ' Desc: Devuelve un array list con los campos del access
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function DataSetCamposLocalAccess(ByVal pstrNombreTabla As String) As DataSet

        Dim lobjCommand As New OleDbCommand(String.Format("SELECT * FROM {0}", pstrNombreTabla), mobjCnnAccess)
        Dim lobjDataAdapter As New OleDbDataAdapter(lobjCommand)
        Dim lobjDataSet As New DataSet
        lobjDataAdapter.FillSchema(lobjDataSet, SchemaType.Source)

        Return lobjDataSet

    End Function

    ' ************************************************************************
    ' DataSetCamposLocalServer
    ' Desc: Devuelve un array list con los campos del server
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function DataSetCamposLocalServer(ByVal pstrNombreTabla As String) As DataSet

        Dim lobjCommand As New SqlCommand(String.Format("SELECT * FROM {0}", pstrNombreTabla), mobjCnnSqlServer)
        Dim lobjDataAdapter As New SqlDataAdapter(lobjCommand)
        Dim lobjDataSet As New DataSet
        lobjDataAdapter.FillSchema(lobjDataSet, SchemaType.Source)

        Return lobjDataSet

    End Function

    ' ************************************************************************
    ' DataSetCamposRemota
    ' Desc: Devuelve un array list con los campos del server
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function DataSetCamposRemota(ByVal pstrNombreTabla As String) As DataSet

        Dim lobjCommand As New OdbcCommand(String.Format("SELECT * FROM {0}", pstrNombreTabla), mobjCnnODBC)
        Dim lobjDataAdapter As New OdbcDataAdapter(lobjCommand)
        Dim lobjDataSet As New DataSet
        lobjDataAdapter.FillSchema(lobjDataSet, SchemaType.Source)

        Return lobjDataSet

    End Function

    ' ************************************************************************
    ' CargarTablas
    ' Desc: Rutina que cierra la conexión a la BBDD
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Sub CargarTablas()

        If Me.TipoBBDD = BBDD.Local Then
            If Me.mobjConfig.TipoLocal = 0 Then
                BuscarTablas(mobjCnnAccess.GetSchema("TABLES"))
            Else
                BuscarTablas(mobjCnnSqlServer.GetSchema("TABLES"))
            End If
        Else
            BuscarTablas(mobjCnnODBC.GetSchema("TABLES"))
        End If

    End Sub

    ' ************************************************************************
    ' CargarTablasAccess
    ' Desc: Rutina que carga las tablas de la BBDD Access
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Sub BuscarTablas(ByVal dt As DataTable)

        Dim larrlstTablas As New ArrayList

        If dt.Rows.Count <> 0 Then
            For lintContador As Integer = 0 To dt.Rows.Count - 1
                If Not IsDBNull(dt.Rows(lintContador).Item("TABLE_NAME")) Then
                    larrlstTablas.Add(dt.Rows(lintContador).Item("TABLE_NAME").ToString())
                    'MessageBox.Show(col.ColumnName + " " + dt.Rows(lintContador).Item(col.ColumnName).ToString())
                End If
            Next
        End If

        If larrlstTablas.Count > 0 Then
            larrlstTablas.Sort()
            CargarArrayTablas(larrlstTablas)
        End If

    End Sub

    ' ************************************************************************
    ' CargarArrayTablas
    ' Desc: Rutina que carga los arrays a los combos de tablas
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Sub CargarArrayTablas(ByVal parrlstTablas As ArrayList)

        Me.cboTablaBioquimica.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaMuestra.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaMicro.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaMicroMuestra.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaPerfilBio.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaPerfilMicro.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaMedicos.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaHC.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaServicio.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaCorrelaciones.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaOrigenes.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaDestinos.Items.AddRange(parrlstTablas.ToArray())
        Me.cboTablaDiagnostico.Items.AddRange(parrlstTablas.ToArray())

    End Sub

    ' ************************************************************************
    ' CerrarConexion
    ' Desc: Rutina que cierra la conexión a la BBDD
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function CerrarConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota
            If Me.TipoBBDD = BBDD.Local Then
                If Me.mobjConfig.TipoLocal = 0 Then
                    If mobjCnnAccess.State <> ConnectionState.Closed Then mobjCnnAccess.Close()
                Else
                    If mobjCnnSqlServer.State <> ConnectionState.Closed Then mobjCnnSqlServer.Close()
                End If
            Else
                If mobjCnnODBC.State <> ConnectionState.Closed Then mobjCnnODBC.Close()
            End If
            Return True
        Catch ex As Exception
            'MessageBox.Show("Error al cerrar la conexion a BBDD" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function
    ' ************************************************************************
    ' AbrirConexion
    ' Desc: Rutina que abre la conexión a la BBDD
    ' NBL: 23/01/2007
    ' ************************************************************************
    Private Function AbrirConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota
            If Me.TipoBBDD = BBDD.Local Then
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' La BBDD es Access
                    mobjCnnAccess = New OleDbConnection(Me.mobjConfig.cnLocal)
                    mobjCnnAccess.Open()
                Else
                    ' SQL Server
                    mobjCnnSqlServer = New SqlConnection(Me.mobjConfig.cnLocal)
                    mobjCnnSqlServer.Open()
                End If
            Else
                ' ODBC hacia OMEGA
                mobjCnnODBC = New OdbcConnection("DSN=" + Me.mobjConfig.DSNRemota)
                mobjCnnODBC.Open()
            End If
            Return True
        Catch ex As Exception
            ' MessageBox.Show("No se ha podido conectar a la BBDD" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ************************************************************************
    ' CargarConfiguracion
    ' Desc: Rutina que carga la configuración de BBDD
    ' NBL: 11/01/2007
    ' ************************************************************************
    Private Sub CargarConfiguracion()

        'Leer un archivo XML y cargarlo en un objeto
        Dim xmlReader As New XmlTextReader(mstrNombreArchivoConfigBBDD)

        'Crear un objeto para deserializar el archivo XML
        Dim Reader As New Serialization.XmlSerializer(GetType(clsConfigBBDD))

        'Deserialziar el archivo xml y cargarlo en un objeto
        mobjConfigBBDD = Reader.Deserialize(xmlReader)

        ' Cargamos en los textboxes los valores
        CargarDatos(mobjConfigBBDD, True)

        'Cerrar Archivo XML
        xmlReader.Close()

    End Sub

    ' ************************************************************************
    ' CargarDatos
    ' Desc: Rutina que carga los textboxes con los valores de la instancia del archivo xml
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarDatos(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextBox As Boolean)

        CargarBioquimica(pobjConfigBBDD, pbolTextBox)
        CargarMuestra(pobjConfigBBDD, pbolTextBox)
        CargarMicrobiologia(pobjConfigBBDD, pbolTextBox)
        CargarMicroMuestra(pobjConfigBBDD, pbolTextBox)
        CargarPerfilesBio(pobjConfigBBDD, pbolTextBox)
        CargarPerfilesMicro(pobjConfigBBDD, pbolTextBox)
        CargarMedicos(pobjConfigBBDD, pbolTextBox)
        CargarHistoriasClinicas(pobjConfigBBDD, pbolTextBox)
        CargarServicios(pobjConfigBBDD, pbolTextBox)
        CargarOrigenes(pobjConfigBBDD, pbolTextBox)
        CargarDestinos(pobjConfigBBDD, pbolTextBox)
        CargarCorrelaciones(pobjConfigBBDD, pbolTextBox)
        CargarDiagnosticos(pobjConfigBBDD, pbolTextBox)

    End Sub

    ' ************************************************************************
    ' CargarDestinos
    ' Desc: Rutina que carga la tabla de destinos
    ' NBL: 16/01/2007
    ' ************************************************************************
    Private Sub CargarDestinos(ByVal pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Destinos
            If pbolTextbox Then
                Me.cboTablaDestinos.Text = .TABLA
                Me.cboCodigoDestinos.Text = .CODIGO
                Me.cboNombreDestinos.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaDestinos.Text.Trim
                .CODIGO = Me.cboCodigoDestinos.Text.Trim
                .NOMBRE = Me.cboNombreDestinos.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarOrigenes
    ' Desc: Rutina que carga la tabla de origenes
    ' NBL: 16/01/2007
    ' ************************************************************************
    Private Sub CargarOrigenes(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Origenes
            If pbolTextbox Then
                Me.cboTablaOrigenes.Text = .TABLA
                Me.cboCodigoOrigen.Text = .CODIGO
                Me.cboNombreOrigen.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaOrigenes.Text.Trim
                .CODIGO = Me.cboCodigoOrigen.Text.Trim
                .NOMBRE = Me.cboNombreOrigen.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarCorrelaciones
    ' Desc: Rutina que carga la tabla de correlaciones
    ' NBL: 16/01/2007
    ' ************************************************************************
    Private Sub CargarCorrelaciones(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Correlaciones
            If pbolTextbox Then
                Me.cboTablaCorrelaciones.Text = .TABLA
                Me.cboIDCorrelacion.Text = .ID
                Me.cboCodigoCampoCorrelacion.Text = .CODIGO_CAMPO
                Me.cboCodigoDesencadenanteCorrelacion.Text = .CODIGO_DESENCADENANTE
                Me.cboCodigoDoctorCorrelacion.Text = .CODIGO_DOCTOR
                Me.cboCodigoServicioCorrelacion.Text = .CODIGO_SERVICIO
                Me.cboCodigoOrigenCorrelacion.Text = .CODIGO_ORIGEN
                Me.cboCodigoDestinoCorrelacion.Text = .CODIGO_DESTINO
                Me.cboCodigoMotivoCorrelacion.Text = .CODIGO_MOTIVO
                Me.cboCodigoTipoCorrelacion.Text = .CODIGO_TIPO
                Me.cboCodigoGrupoFacturacionCorrelacion.Text = .CODIGO_GRUPO_FACTURACION
                Me.cboCodigoCargoCorrelacion.Text = .CODIGO_CARGO
            Else
                .TABLA = Me.cboTablaCorrelaciones.Text.Trim
                .ID = Me.cboIDCorrelacion.Text
                .CODIGO_CAMPO = Me.cboCodigoCampoCorrelacion.Text
                .CODIGO_DESENCADENANTE = Me.cboCodigoDesencadenanteCorrelacion.Text
                .CODIGO_DOCTOR = Me.cboCodigoDoctorCorrelacion.Text
                .CODIGO_SERVICIO = Me.cboCodigoServicioCorrelacion.Text
                .CODIGO_ORIGEN = Me.cboCodigoOrigenCorrelacion.Text
                .CODIGO_DESTINO = Me.cboCodigoDestinoCorrelacion.Text
                .CODIGO_MOTIVO = Me.cboCodigoMotivoCorrelacion.Text
                .CODIGO_TIPO = Me.cboCodigoTipoCorrelacion.Text
                .CODIGO_GRUPO_FACTURACION = Me.cboCodigoGrupoFacturacionCorrelacion.Text
                .CODIGO_CARGO = Me.cboCodigoCargoCorrelacion.Text
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarServicios
    ' Desc: Rutina que carga la tabla de servicios
    ' NBL: 16/01/2007
    ' ************************************************************************
    Private Sub CargarServicios(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Servicios
            If pbolTextbox Then
                Me.cboTablaServicio.Text = .TABLA
                Me.cboCodigoServicio.Text = .CODIGO
                Me.cboNombreServicio.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaServicio.Text.Trim
                .CODIGO = Me.cboCodigoServicio.Text.Trim
                .NOMBRE = Me.cboNombreServicio.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarDiagnosticos
    ' Desc: Rutina que carga la tabla de diagnósticos
    ' NBL: 11/06/2007
    ' ************************************************************************
    Private Sub CargarDiagnosticos(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Diagnostico
            If pbolTextbox Then
                Me.cboTablaDiagnostico.Text = .TABLA
                Me.cboCodigoDiagnostico.Text = .CODIGO
                Me.cboNombreDiagnostico.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaDiagnostico.Text.Trim
                .CODIGO = Me.cboCodigoDiagnostico.Text.Trim
                .NOMBRE = Me.cboNombreDiagnostico.Text.Trim
            End If
        End With

    End Sub


    ' ************************************************************************
    ' CargarHistoriasClinicas
    ' Desc: Rutina que carga la tabla de historias clínicas
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarHistoriasClinicas(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.HistoriasClinicas
            If pbolTextbox Then
                Me.cboTablaHC.Text = .TABLA
                Me.cboIDHC.Text = .ID
                Me.cboNumHC.Text = .NUM_HISTORIA
                Me.cboApellidosHC.Text = .APELLIDOS
                Me.cboNombreHC.Text = .NOMBRE
                Me.cboNumeroSSHC.Text = .NUM_SS
                Me.cboCodSexoHC.Text = .COD_SEXO
                Me.cboFechaNacimientoHC.Text = .FECHA_NACIMIENTO
                Me.cboDNIHC.Text = .DNI
                Me.cboDireccionHC.Text = .DIRECCION
                Me.cboPoblacionHC.Text = .POBLACION
                Me.cboProvinciaHC.Text = .COD_PROVINCIA
                Me.cboCPHC.Text = .COD_POSTAL
                Me.cboTelefonoHC.Text = .TELEFONO
                Me.cboNombreCompletoHC.Text = .NOMBRE_COMPLETO
            Else
                .TABLA = Me.cboTablaHC.Text.Trim
                .ID = Me.cboIDHC.Text.Trim
                .NUM_HISTORIA = Me.cboNumHC.Text.Trim
                .APELLIDOS = Me.cboApellidosHC.Text.Trim
                .NOMBRE = Me.cboNombreHC.Text.Trim
                .NUM_SS = Me.cboNumeroSSHC.Text.Trim
                .COD_SEXO = Me.cboCodSexoHC.Text.Trim
                .FECHA_NACIMIENTO = Me.cboFechaNacimientoHC.Text.Trim
                .DNI = Me.cboDNIHC.Text.Trim
                .DIRECCION = Me.cboDireccionHC.Text.Trim
                .POBLACION = Me.cboPoblacionHC.Text.Trim
                .COD_PROVINCIA = Me.cboProvinciaHC.Text.Trim
                .COD_POSTAL = Me.cboCPHC.Text.Trim
                .TELEFONO = Me.cboTelefonoHC.Text.Trim
                .NOMBRE_COMPLETO = Me.cboNombreCompletoHC.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarMedicos
    ' Desc: Rutina que carga los textboxes de perfiles médicos
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarMedicos(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Medicos
            If pbolTextbox Then
                Me.cboTablaMedicos.Text = .TABLA
                Me.cboCodigoMedico.Text = .CODIGO
                Me.cboNombreMedico.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaMedicos.Text.Trim
                .CODIGO = Me.cboCodigoMedico.Text.Trim
                .NOMBRE = Me.cboNombreMedico.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarPerfilesMicro
    ' Desc: Rutina que carga los textboxes de perfiles microbiología
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarPerfilesMicro(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.PerfilMicro
            If pbolTextbox Then
                Me.cboTablaPerfilMicro.Text = .TABLA
                Me.cboCodigoPerfilMicro.Text = .CODIGO
                Me.cboNombrePerfilMicro.Text = .NOMBRE
                Me.cboCodMuestraPerfilMicro.Text = .CODIGO_MUESTRA
            Else
                .TABLA = Me.cboTablaPerfilMicro.Text.Trim
                .CODIGO = Me.cboCodigoPerfilMicro.Text.Trim
                .NOMBRE = Me.cboNombrePerfilMicro.Text.Trim
                .CODIGO_MUESTRA = Me.cboCodMuestraPerfilMicro.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarPerfilesBio
    ' Desc: Rutina que carga los textboxes de perfiles bioquímica
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarPerfilesBio(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.PerfilBioquimica
            If pbolTextbox Then
                Me.cboTablaPerfilBio.Text = .TABLA
                Me.cboCodigoPerfilBio.Text = .CODIGO
                Me.cboNombrePerfilBio.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaPerfilBio.Text.Trim
                .CODIGO = Me.cboCodigoPerfilBio.Text.Trim
                .NOMBRE = Me.cboNombrePerfilBio.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarMicroMuestra
    ' Desc: Rutina que carga los textboxes de micro-muestra
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarMicroMuestra(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.MicroMuestra
            If pbolTextbox Then
                Me.cboTablaMicroMuestra.Text = .TABLA
                Me.cboCodPrueba.Text = .CODIGO_PRUEBA
                Me.cboCodMuestra.Text = .CODIGO_MUESTRA
            Else
                .TABLA = Me.cboTablaMicroMuestra.Text.Trim
                .CODIGO_PRUEBA = Me.cboCodPrueba.Text.Trim
                .CODIGO_MUESTRA = Me.cboCodMuestra.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarMicrobiologia
    ' Desc: Rutina que carga los textboxes de microbiología
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarMicrobiologia(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Microbiologia
            If pbolTextbox Then
                Me.cboTablaMicro.Text = .TABLA
                Me.cboCodigoMicro.Text = .CODIGO
                Me.cboAbrvMicro.Text = .ABRV
                Me.cboNombreMicro.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaMicro.Text.Trim
                .CODIGO = Me.cboCodigoMicro.Text.Trim
                .ABRV = Me.cboAbrvMicro.Text.Trim
                .NOMBRE = Me.cboNombreMicro.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarMuestra
    ' Desc: Rutina que carga los textboxes de muestras
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarMuestra(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Muestra
            If pbolTextbox Then
                Me.cboTablaMuestra.Text = .TABLA
                Me.cboCodigoMuestra.Text = .CODIGO
                Me.cboAbrvMuestra.Text = .ABRV
                Me.cboNombreMuestra.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaMuestra.Text.Trim
                .CODIGO = Me.cboCodigoMuestra.Text.Trim
                .ABRV = Me.cboAbrvMuestra.Text.Trim
                .NOMBRE = Me.cboNombreMuestra.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' CargarBioquimica
    ' Desc: Rutina que carga los textboxes de bioquímica
    ' NBL: 15/01/2007
    ' ************************************************************************
    Private Sub CargarBioquimica(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pbolTextbox As Boolean)

        With pobjConfigBBDD.Bioquimica
            If pbolTextbox Then
                Me.cboTablaBioquimica.Text = .TABLA
                Me.cboCodigoBioquimica.Text = .CODIGO
                Me.cboAbrvBioquimica.Text = .ABRV
                Me.cboNombreBioquimica.Text = .NOMBRE
            Else
                .TABLA = Me.cboTablaBioquimica.Text.Trim
                .CODIGO = Me.cboCodigoBioquimica.Text.Trim
                .ABRV = Me.cboAbrvBioquimica.Text.Trim
                .NOMBRE = Me.cboNombreBioquimica.Text.Trim
            End If
        End With

    End Sub

    ' ************************************************************************
    ' GuardarConfiguracion
    ' Desc: Rutina que guarda la configuración de BBDD
    ' NBL: 11/01/2007
    ' ************************************************************************
    Private Sub GuardarConfiguracion()

        ' En primer lugar miramos si existe el archivo de configuración y lo borramos.
        If My.Computer.FileSystem.FileExists(mstrNombreArchivoConfigBBDD) Then _
            My.Computer.FileSystem.DeleteFile(mstrNombreArchivoConfigBBDD)

        If mobjConfigBBDD Is Nothing Then mobjConfigBBDD = New clsConfigBBDD
        CargarDatos(mobjConfigBBDD, False)

        'Crear un objeto serializado para la clase contactos
        Dim objWriter As New Serialization.XmlSerializer(GetType(clsConfigBBDD))
        'Crear un objeto file de tipo StremWriter para almacenar el documento xml
        Dim objFile As New StreamWriter(mstrNombreArchivoConfigBBDD)
        'Serializar y crear el documento XML
        objWriter.Serialize(objFile, mobjConfigBBDD)
        'Cerrar el archivo
        objFile.Close()

    End Sub

    Private Sub DialogoConfiguracionBBDD_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If MessageBox.Show("¿Guardar la configuración?", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
            GuardarConfiguracion()
        End If

        If Me.TipoBBDD = BBDD.Local Then

            If Me.Top > 0 Then
                My.Settings.DialogoConfiguracionBBDDLocalLocation = Me.Location
            End If
            If Me.WindowState <> FormWindowState.Maximized Then My.Settings.DialogoConfiguracionBBDDLocalSize = Me.Size
            My.Settings.DialogoConfiguracionBBDDLocalState = Me.WindowState

        Else

            If Me.Top > 0 Then
                My.Settings.DialogoConfiguracionBBDDRemotaLocation = Me.Location
            End If

            If Me.WindowState <> FormWindowState.Maximized Then My.Settings.DialogoConfiguracionBBDDRemotaSize = Me.Size
            My.Settings.DialogoConfiguracionBBDDRemotaState = Me.WindowState

        End If

        My.Settings.Save()

        ' Finalmente cerramos la conexión
        CerrarConexion()

    End Sub

    Private Sub DialogoConfiguracionBBDD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Me.TipoBBDD = BBDD.Local Then
            If Not My.Settings.DialogoConfiguracionBBDDLocalLocation.IsEmpty Then
                Me.Location = My.Settings.DialogoConfiguracionBBDDLocalLocation
                If Me.Top < 0 Then
                    Me.Top = 0
                    Me.Left = 0
                End If
            End If
            Me.Size = My.Settings.DialogoConfiguracionBBDDLocalSize
            If Me.Width < 913 Then Me.Width = 913
            If Me.Height < 567 Then Me.Height = 567
            If My.Settings.DialogoConfiguracionBBDDLocalState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Maximized
            Else
                Me.WindowState = FormWindowState.Normal
            End If
        Else

            If Not My.Settings.DialogoConfiguracionBBDDRemotaLocation.IsEmpty Then
                Me.Location = My.Settings.DialogoConfiguracionBBDDRemotaLocation
                If Me.Top < 0 Then
                    Me.Top = 0
                    Me.Left = 0
                End If
            End If
            Me.Size = My.Settings.DialogoConfiguracionBBDDRemotaSize
            If Me.Width < 913 Then Me.Width = 913
            If Me.Height < 567 Then Me.Height = 567
            If My.Settings.DialogoConfiguracionBBDDRemotaState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Maximized
            Else
                Me.WindowState = FormWindowState.Normal
            End If

        End If

    End Sub

    Private Sub cboTablaBioquimica_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaBioquimica.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaBioquimica.Text, Me.tabBioquimica, _
                                                                                            "cboCodigoBioquimica", "cboAbrvBioquimica", "cboNombreBioquimica")
    End Sub

    Private Sub cboCodigoMuestra_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboCodigoMuestra.SelectedIndexChanged

    End Sub

    Private Sub cboDNIHC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDNIHC.SelectedIndexChanged

    End Sub

    Private Sub cboTelefonoHC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTelefonoHC.SelectedIndexChanged

    End Sub

    Private Sub cboTablaMuestra_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTablaMuestra.SelectedIndexChanged

    End Sub

    Private Sub cboTablaMuestra_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaMuestra.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaMuestra.Text, Me.tabMuestra, _
                                                                        "cboCodigoMuestra", "cboAbrvMuestra", "cboNombreMuestra")
    End Sub

    Private Sub cboTablaMicro_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaMicro.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaMicro.Text, Me.tabMicrobiologia, _
                                                                                "cboCodigoMicro", "cboAbrvMicro", "cboNombreMicro")
    End Sub

    Private Sub cboTablaMicroMuestra_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTablaMicroMuestra.SelectedIndexChanged
        
    End Sub

    Private Sub cboTablaPerfilBio_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTablaPerfilBio.SelectedIndexChanged

    End Sub

    Private Sub cboTablaPerfilBio_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaPerfilBio.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaPerfilBio.Text, Me.tabPerfilBioquimica, _
                                                                                    "cboCodigoPerfilBio", "cboNombrePerfilBio")
    End Sub

    Private Sub cboTablaMicroMuestra_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaMicroMuestra.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaMicroMuestra.Text, Me.tabMicroMuestra, _
                                                                                            "cboCodPrueba", "cboCodMuestra")
    End Sub

    Private Sub cboTablaPerfilMicro_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaPerfilMicro.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaPerfilMicro.Text, Me.tabPerfilMicro, _
                                                                                       "cboCodigoPerfilMicro", "cboNombrePerfilMicro", "cboCodMuestraPerfilMicro")
    End Sub

    Private Sub cboTablaMedicos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaMedicos.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaMedicos.Text, Me.tabMedicos, _
                                                                                    "cboCodigoMedico", "cboNombreMedico")
    End Sub

    Private Sub cboTablaHC_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaHC.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaHC.Text, Me.tabHC, _
                                                            "cboIDHC", "cboNumHC", "cboApellidosHC", "cboNombreHC", "cboNumeroSSHC", "cboCodSexoHC", _
                                                            "cboFechaNacimientoHC", "cboDNIHC", "cboDireccionHC", "cboPoblacionHC", "cboProvinciaHC", "cboCPHC", _
                                                            "cboTelefonoHC", "cboNombreCompletoHC")
    End Sub

    Private Sub cboTablaServicio_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaServicio.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaServicio.Text, Me.tabServicios, _
                                                                                    "cboCodigoServicio", "cboNombreServicio")
    End Sub

    Private Sub cboTablaCorrelaciones_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaCorrelaciones.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaCorrelaciones.Text, Me.tabCorrelaciones, _
                                                                                "cboIDCorrelacion", "cboCodigoCampoCorrelacion", "cboCodigoDesencadenanteCorrelacion", _
                                                                                "cboCodigoDoctorCorrelacion", "cboCodigoServicioCorrelacion", "cboCodigoOrigenCorrelacion", _
                                                                                "cboCodigoDestinoCorrelacion", "cboCodigoMotivoCorrelacion", "cboCodigoTipoCorrelacion", _
                                                                                "cboCodigoGrupoFacturacionCorrelacion", "cboCodigoCargoCorrelacion")
    End Sub

    Private Sub cboTablaOrigenes_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaOrigenes.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaOrigenes.Text, Me.tabOrigenes, _
                                                                                "cboCodigoOrigen", "cboNombreOrigen")
    End Sub

    Private Sub cboTablaDestinos_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaDestinos.Validating
        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaDestinos.Text, Me.tabDestinos, _
                                                                             "cboCodigoDestinos", "cboNombreDestinos")
    End Sub

    Private Sub cboTablaBioquimica_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTablaBioquimica.SelectedIndexChanged

    End Sub

    Private Sub cboTablaDestinos_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboTablaDestinos.SelectedIndexChanged

    End Sub

    Private Sub cboTablaDiagnostico_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles cboTablaDiagnostico.Validating

        If ConexionActiva() Then CargarNombresCampos(Me.cboTablaDiagnostico.Text, Me.tabDiagnosticos, _
                                                                     "cboCodigoDiagnostico", "cboNombreDiagnostico")
    End Sub
End Class