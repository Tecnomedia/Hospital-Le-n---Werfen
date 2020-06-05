Imports System.Xml
Imports System.IO
Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Text

' PERFIL --> ROJO
' PRUEBA --> VERDE

Friend Class DialogoPruebas

    ' Resultado del diálogo
    Public Resultado As String
    ' Guardamos en esta variable la selección de entrada al diálogo por si cancelan
    Private pstrSeleccionEntrada As String
    ' Variable chivata con la que controlamos que aceptamos la selección que hay
    Private mbolAceptandoDialogo As Boolean = False

    ' Instancias de configuración
    Private mobjConfig As clsConfig
    Private mobjConfigBBDD As clsConfigBBDD

    ' Ponemos las tres tipos de conexiones que puede haber en este formulario
    Private mobjCnnAccess As OleDbConnection
    Private mobjCnnSqlServer As SqlConnection
    Private mobjCnnODBC As OdbcConnection

    Sub New(ByVal pstrTituloDialogo As String, ByVal pstrPruebasPreseleccionadas As String)

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        Me.Text = pstrTituloDialogo

        ' Add any initialization after the InitializeComponent() call.
        Dim mstrNombreArchivoConfig As String = clsUtil.DLLPath(True) + "Config.xml"
        clsUtil.CargarConfiguracion(mobjConfig, mstrNombreArchivoConfig)

        Dim mstrNombreArchivoConfigBBDD As String = ""
        If Me.mobjConfig.Conexion = BBDD.Local Then
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigLocalBBDD.xml"
        Else
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigRemotaBBDD.xml"
        End If
        clsUtil.CargarConfiguracionBBDD(mobjConfigBBDD, mstrNombreArchivoConfigBBDD)

        ' Cargamos las pruebas seleccionadas
        CargarPruebasSeleccionadas(pstrPruebasPreseleccionadas)

        Me.pstrSeleccionEntrada = pstrPruebasPreseleccionadas

    End Sub

    ' ********************************************************************************
    ' getDescripcionPruebas
    ' Desc: Función que devuelve la abrv y la descripción de una prueba
    ' NBL: 25/6/2009
    ' NBL: 7/05/2010 Introduzco un nuevo parámetro opcional, que nos indica si hemos de buscar solo en micro y otro para decidir si es perfil o no
    ' ********************************************************************************
    Public Sub getAbrvDescripcionPruebas(ByRef pstrCodigo As String, ByRef pstrAbrv As String, ByRef pstrDescripcion As String, _
                                                                ByRef pstrTipoPrueba As String, ByVal pobjCnnAccess As OleDbConnection, _
                                                                ByVal pobjCnnSqlServer As SqlConnection, ByVal pobjCnnODBC As OdbcConnection, Optional ByVal pbolMicro As Boolean = False, Optional ByVal pbolPerfil As Boolean = False)

        Me.mobjCnnAccess = pobjCnnAccess
        Me.mobjCnnSqlServer = pobjCnnSqlServer
        Me.mobjCnnODBC = pobjCnnODBC

        If Me.mobjConfig.Conexion = BBDD.Local Then
            ' BBDD local
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access
                getDescripcionPruebaPreseleccionadaLocalAccess(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                'If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalAccess(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
            Else
                ' SQL Server
                getDescripcionPruebaPreseleccionadaLocalServer(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                ' If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalServer(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
            End If
        Else
            ' BBDD remota
            getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba, pbolMicro, pbolPerfil)
            ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
            ' If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaRemota(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
        End If

    End Sub

    ' ********************************************************************************
    ' getDescripcionPruebas
    ' Desc: Función que devuelve la abrv y la descripción de una prueba, en esta versión también buscamos la de muestra
    ' NBL: 21/10/2009
    ' NBL: 7/05/2010 Introduzco un nuevo parámetro opcional, que nos indica si hemos de buscar solo en micro y otro para decidir si es perfil o no
    ' ********************************************************************************
    Public Sub getAbrvDescripcionPruebas(ByRef pstrCodigo As String, ByRef pstrAbrv As String, ByRef pstrDescripcion As String, _
                                                                ByRef pstrCodigoMuestra As String, ByRef pstrAbrvMuestra As String, ByRef pstrDescripcionMuestra As String, _
                                                                ByRef pstrTipoPrueba As String, ByVal pobjCnnAccess As OleDbConnection, _
                                                                ByVal pobjCnnSqlServer As SqlConnection, ByVal pobjCnnODBC As OdbcConnection, _
                                                                Optional ByVal pbolMicro As Boolean = False, Optional ByVal pbolPerfil As Boolean = False)

        Me.mobjCnnAccess = pobjCnnAccess
        Me.mobjCnnSqlServer = pobjCnnSqlServer
        Me.mobjCnnODBC = pobjCnnODBC

        If Me.mobjConfig.Conexion = BBDD.Local Then
            ' BBDD local
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access
                getDescripcionPruebaPreseleccionadaLocalAccess(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalAccess(pstrCodigoMuestra, pstrAbrvMuestra, pstrDescripcionMuestra)
            Else
                ' SQL Server
                getDescripcionPruebaPreseleccionadaLocalServer(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalServer(pstrCodigoMuestra, pstrAbrvMuestra, pstrDescripcionMuestra)
            End If
        Else
            ' BBDD remota
            If pstrCodigoMuestra.Length = 0 Then
                getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba, pbolMicro, True)
            Else
                getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, pstrAbrv, pstrDescripcion, pstrTipoPrueba, pbolMicro, pbolPerfil)
            End If
            ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
            If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaRemota(pstrCodigoMuestra, pstrAbrvMuestra, pstrDescripcionMuestra)
        End If

    End Sub

    ' ************************************************************************
    ' CargarPruebasSeleccionadas
    ' Desc: Función que carga las pruebas que ya están seleccionadas y se pasan como parámetro
    ' NBL: 14/02/2007
    ' ************************************************************************
    Private Sub CargarPruebasSeleccionadas(ByVal pstrPruebas As String)

        ' Primero miramos si hay pruebas o no
        If pstrPruebas.Trim.Length = 0 Then Exit Sub

        ' Hemos de hacer un split de comas para separar las pruebas
        Dim lstrPruebas() As String = pstrPruebas.Split(",")

        ' Lo que hay pasado no es una prueba, por lo menos tendría que haber longitud 2
        If pstrPruebas.Length <= 1 Then Exit Sub

        If Not AbrirConexion() Then Exit Sub

        For lintContador As Integer = 1 To lstrPruebas.Length - 1
            ' Tenemos un bucle por las diferentes pruebas
            Dim lstrPrueba() As String = lstrPruebas(lintContador).Split("|")
            Dim lstrCodigoPrueba() As String = lstrPrueba(0).Split("^")
            Dim lstrCodigo As String = lstrCodigoPrueba(0)
            Dim lstrTipoPrueba As String = lstrCodigoPrueba(1)
            Dim lstrCodigoMuestra As String = lstrPrueba(1)

            AddPruebaPreseleccionada(lstrCodigo, lstrTipoPrueba, lstrCodigoMuestra)
        Next

        CerrarConexion()

    End Sub

    ' ************************************************************************
    ' AddPruebaPreseleccionada
    ' Desc: Rutina que busca la abrv y el nombre de la prueba
    ' NBL: 14/02/2007
    ' ************************************************************************
    Private Sub AddPruebaPreseleccionada(ByVal pstrCodigo As String, ByVal pstrTipoPrueba As String, _
                                                                ByVal pstrCodigoMuestra As String)

        Dim lstrAbrv As String = "", lstrNombre As String = "", lstrPruebaPerfil As String = ""
        Dim lstrAbrvMuestra As String = "", lstrNombreMuestra As String = ""

        If Me.mobjConfig.Conexion = BBDD.Local Then
            ' BBDD local
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access
                getDescripcionPruebaPreseleccionadaLocalAccess(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalAccess(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
            Else
                ' SQL Server
                getDescripcionPruebaPreseleccionadaLocalServer(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil)
                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaLocalServer(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
            End If
        Else
            ' BBDD remota
            If pstrTipoPrueba = "M" Then
                If pstrCodigoMuestra.Trim.Length = 0 Then
                    getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil, True, True)
                Else
                    getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil, True, False)
                End If
            ElseIf pstrTipoPrueba = "B" Then
                getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil, False)
            Else
                getDescripcionPruebaPreseleccionadaRemota(pstrCodigo, pstrTipoPrueba, lstrAbrv, lstrNombre, lstrPruebaPerfil, False)
            End If


                ' Si hay codigo de muestra hay que buscar su abrv y su descripcion
                If pstrCodigoMuestra.Length > 0 Then getDescripcionMuestraPreseleccionadaRemota(pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra)
            End If

        If lstrPruebaPerfil.Length = 0 And pstrCodigoMuestra.Length = 0 Then lstrPruebaPerfil = "Perfil"

        'OJO, hay que meter aquí lo de la selección de la prueba 
        Me.dgvPruebasTotal.Rows.Add(pstrCodigo, lstrAbrv, lstrNombre, pstrCodigoMuestra, lstrAbrvMuestra, lstrNombreMuestra, pstrTipoPrueba, lstrPruebaPerfil)

    End Sub

    ' ************************************************************************
    ' getDescripcionPruebaPreseleccionadaRemota
    ' Desc: Rutina que busca abrv y nombre a partir de un código dado remota
    ' NBL: 15/02/2007
    ' NBL: 7/05/2010 Introduzco un nuevo parámetro opcional, que nos indica si hemos de buscar solo en micro y otro para decidir si es perfil o no
    ' ************************************************************************
    Public Sub getDescripcionPruebaPreseleccionadaRemota(ByVal pstrCodigo As String, ByVal pstrTipoPrueba As String, _
                                                                                                ByRef pstrAbrv As String, ByRef pstrNombre As String, _
                                                                                                ByRef pstrPruebaPerfil As String, Optional ByVal pbolMicro As Boolean = False, Optional ByVal pbolPerfil As Boolean = False)

        ' PRUEBAS --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) And Not pbolMicro Then
            Dim lobjCommandPrueba As New OdbcCommand(sqlGetPruebasBioquimicaByCodeRemota(pstrCodigo), mobjCnnODBC)
            Dim lobjParamPrueba As New OdbcParameter("@BUSQUEDA", OdbcType.Double)
            lobjParamPrueba.Value = CType(pstrCodigo, Integer)
            lobjCommandPrueba.Parameters.Add(lobjParamPrueba)
            Dim lobjDataReaderPrueba As OdbcDataReader = lobjCommandPrueba.ExecuteReader()

            If lobjDataReaderPrueba.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPrueba.IsDBNull(1) Then
                    pstrAbrv = ""
                Else
                    pstrAbrv = lobjDataReaderPrueba.GetString(1)
                End If
                ' Nombre
                If lobjDataReaderPrueba.IsDBNull(2) Then
                    pstrNombre = ""
                Else
                    pstrNombre = lobjDataReaderPrueba.GetString(2)
                End If
                pstrPruebaPerfil = "Prueba"
                lobjDataReaderPrueba.Close()
                Exit Sub
            End If
        End If

        ' PERFILES --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) And Not pbolMicro Then
            Dim lobjCommandPerfiles As New OdbcCommand(sqlGetPerfilesBioquimicaByCodeRemota(pstrCodigo), mobjCnnODBC)
            Dim lobjParamPerfiles As New OdbcParameter("@BUSQUEDA", OdbcType.Double)
            lobjParamPerfiles.Value = CType(pstrCodigo, Integer)
            lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
            Dim lobjDataReaderPerfiles As OdbcDataReader = lobjCommandPerfiles.ExecuteReader()

            If lobjDataReaderPerfiles.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPerfiles.IsDBNull(1) Then
                    pstrAbrv = pstrCodigo
                    pstrNombre = ""
                Else
                    pstrAbrv = pstrCodigo
                    pstrNombre = lobjDataReaderPerfiles.GetString(1)
                End If
                pstrPruebaPerfil = "Perfil"
                lobjDataReaderPerfiles.Close()
                Exit Sub
            End If
        End If

        ' MICROBIOLOGIA -----------------------------------------------------------------------------------------------------------
        Dim lobjCommandMicro As New OdbcCommand(sqlGetMicroByCodeRemotaPreseleccionado(pstrCodigo, pbolPerfil), mobjCnnODBC)
        'Dim lobjParamMicro As New OdbcParameter("@BUSQUEDA", OdbcType.Double)
        'lobjParamMicro.Value = CType(pstrCodigo, Integer)
        'lobjCommandMicro.Parameters.Add(lobjParamMicro)
        Dim lobjDataReaderMicro As OdbcDataReader = lobjCommandMicro.ExecuteReader()

        If lobjDataReaderMicro.Read() Then
            ' Si hay datos, pillamos la abreviación y el nombre
            ' Abreviación
            If lobjDataReaderMicro.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMicro.GetString(1)
            End If
            ' Nombre
            If lobjDataReaderMicro.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMicro.GetString(2)
            End If
            If IsNumeric(pstrCodigo) Then
                pstrPruebaPerfil = "Prueba"
            Else
                pstrPruebaPerfil = "Perfil"
            End If
            lobjDataReaderMicro.Close()
            Exit Sub
        End If

        pstrAbrv = ""
        pstrNombre = ""

    End Sub

    ' ************************************************************************
    ' getDescripcionPruebaPreseleccionadaLocalServer
    ' Desc: Rutina que busca abrv y nombre a partir de un código dado
    ' NBL: 15/02/2007
    ' ************************************************************************
    Public Sub getDescripcionPruebaPreseleccionadaLocalServer(ByVal pstrCodigo As String, ByVal pstrTipoPrueba As String, _
                                                                                                ByRef pstrAbrv As String, ByRef pstrNombre As String, _
                                                                                                ByRef pstrPruebaPerfil As String)

        ' PRUEBAS --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) Then
            Dim lobjCommandPrueba As New SqlCommand(sqlGetPruebasBioquimicaByCodeLocal(), mobjCnnSqlServer)
            Dim lobjParamPrueba As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
            lobjParamPrueba.Value = CType(pstrCodigo, Integer)
            lobjCommandPrueba.Parameters.Add(lobjParamPrueba)
            Dim lobjDataReaderPrueba As SqlDataReader = lobjCommandPrueba.ExecuteReader()

            If lobjDataReaderPrueba.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPrueba.IsDBNull(1) Then
                    pstrAbrv = ""
                Else
                    pstrAbrv = lobjDataReaderPrueba.GetString(1)
                End If
                ' Nombre
                If lobjDataReaderPrueba.IsDBNull(2) Then
                    pstrNombre = ""
                Else
                    pstrNombre = lobjDataReaderPrueba.GetString(2)
                End If
                pstrPruebaPerfil = "Prueba"
                lobjDataReaderPrueba.Close()
                Exit Sub
            End If
            lobjDataReaderPrueba.Close()
        End If

        ' PERFILES --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) Then
            Dim lobjCommandPerfiles As New SqlCommand(sqlGetPerfilesBioquimicaByCodeLocal(), mobjCnnSqlServer)
            Dim lobjParamPerfiles As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
            lobjParamPerfiles.Value = CType(pstrCodigo, Integer)
            lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
            Dim lobjDataReaderPerfiles As SqlDataReader = lobjCommandPerfiles.ExecuteReader()

            If lobjDataReaderPerfiles.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPerfiles.IsDBNull(1) Then
                    pstrAbrv = pstrCodigo
                    pstrNombre = ""
                Else
                    pstrAbrv = pstrCodigo
                    pstrNombre = lobjDataReaderPerfiles.GetString(1)
                End If
                pstrPruebaPerfil = "Perfil"
                lobjDataReaderPerfiles.Close()
                Exit Sub
            End If
            lobjDataReaderPerfiles.Close()
        End If

        ' MICROBIOLOGIA -----------------------------------------------------------------------------------------------------------
        Dim lobjCommandMicro As New SqlCommand(sqlGetMicroByCodeLocalPreseleccionado(IsNumeric(pstrCodigo)), mobjCnnSqlServer)
        If IsNumeric(pstrCodigo) Then
            Dim lobjParamMicro As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
            lobjParamMicro.Value = CType(pstrCodigo, Integer)
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        Else
            Dim lobjParamMicro As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
            lobjParamMicro.Value = pstrCodigo
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        End If
        Dim lobjDataReaderMicro As SqlDataReader = lobjCommandMicro.ExecuteReader()

        If lobjDataReaderMicro.Read() Then
            ' Si hay datos, pillamos la abreviación y el nombre
            ' Abreviación
            If lobjDataReaderMicro.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMicro.GetString(1)
            End If
            ' Nombre
            If lobjDataReaderMicro.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMicro.GetString(2)
            End If
            lobjDataReaderMicro.Close()
            If IsNumeric(pstrCodigo) Then
                pstrPruebaPerfil = "Prueba"
            Else
                pstrPruebaPerfil = "Perfil"
            End If
            Exit Sub
        End If

        lobjDataReaderMicro.Close()

        pstrAbrv = ""
        pstrNombre = ""

    End Sub

    ' ************************************************************************
    ' getDescripcionMuestraPreseleccionadaLocalAccess
    ' Desc: Función 
    ' NBL: 15/02/2007
    ' ************************************************************************
    Private Sub getDescripcionMuestraPreseleccionadaLocalAccess(ByVal pstrCodigo As String, ByRef pstrAbrv As String, _
                                                                                                        ByRef pstrNombre As String)

        Dim lobjCommandMuestras As New OleDbCommand(sqlGetMuestrasByCodeLocal(), mobjCnnAccess)
        Dim lobjParamMuestras As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
        lobjParamMuestras.Value = CType(pstrCodigo, Integer)
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As OleDbDataReader = lobjCommandMuestras.ExecuteReader()

        If lobjDataReaderMuestras.Read() Then

            If lobjDataReaderMuestras.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMuestras.GetString(1)
            End If

            If lobjDataReaderMuestras.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMuestras.GetString(2)
            End If

        End If

        lobjDataReaderMuestras.Close()

    End Sub

    ' ************************************************************************
    ' getDescripcionMuestraPreseleccionadaLocalServer
    ' Desc: Función 
    ' NBL: 15/02/2007
    ' ************************************************************************
    Private Sub getDescripcionMuestraPreseleccionadaLocalServer(ByVal pstrCodigo As String, ByRef pstrAbrv As String, _
                                                                                                        ByRef pstrNombre As String)

        Dim lobjCommandMuestras As New SqlCommand(sqlGetMuestrasByCodeLocal(), mobjCnnSqlServer)
        Dim lobjParamMuestras As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
        lobjParamMuestras.Value = CType(pstrCodigo, Integer)
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As SqlDataReader = lobjCommandMuestras.ExecuteReader()

        If lobjDataReaderMuestras.Read() Then

            If lobjDataReaderMuestras.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMuestras.GetString(1)
            End If

            If lobjDataReaderMuestras.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMuestras.GetString(2)
            End If

        End If

        lobjDataReaderMuestras.Close()

    End Sub

    ' ************************************************************************
    ' getDescripcionMuestraPreseleccionadaRemota
    ' Desc: Función 
    ' NBL: 15/02/2007
    ' ************************************************************************
    Private Sub getDescripcionMuestraPreseleccionadaRemota(ByVal pstrCodigo As String, ByRef pstrAbrv As String, _
                                                                                                        ByRef pstrNombre As String)

        Dim lobjCommandMuestras As New OdbcCommand(sqlGetMuestrasByCodeRemota(pstrCodigo), mobjCnnODBC)
        Dim lobjDataReaderMuestras As OdbcDataReader = lobjCommandMuestras.ExecuteReader()

        If lobjDataReaderMuestras.Read() Then

            If lobjDataReaderMuestras.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMuestras.GetString(1)
            End If

            If lobjDataReaderMuestras.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMuestras.GetString(2)
            End If

        End If

        lobjDataReaderMuestras.Close()

    End Sub

    ' ************************************************************************
    ' getDescripcionPruebaPreseleccionadaLocalAccess
    ' Desc: Función 
    ' NBL: 15/02/2007
    ' ************************************************************************
    Public Sub getDescripcionPruebaPreseleccionadaLocalAccess(ByVal pstrCodigo As String, ByVal pstrTipoPrueba As String, _
                                                                                                ByRef pstrAbrv As String, ByRef pstrNombre As String, _
                                                                                                ByRef pstrPruebaPerfil As String)

        ' PRUEBAS --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) Then
            Dim lobjCommandPrueba As New OleDbCommand(sqlGetPruebasBioquimicaByCodeLocal(), mobjCnnAccess)
            Dim lobjParamPrueba As New OleDbParameter("@BUSQUEDA", OleDbType.Double)

            lobjParamPrueba.Value = CType(pstrCodigo, Integer)

            lobjCommandPrueba.Parameters.Add(lobjParamPrueba)
            Dim lobjDataReaderPrueba As OleDbDataReader = lobjCommandPrueba.ExecuteReader()

            If lobjDataReaderPrueba.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPrueba.IsDBNull(1) Then
                    pstrAbrv = ""
                Else
                    pstrAbrv = lobjDataReaderPrueba.GetString(1)
                End If
                ' Nombre
                If lobjDataReaderPrueba.IsDBNull(2) Then
                    pstrNombre = ""
                Else
                    pstrNombre = lobjDataReaderPrueba.GetString(2)
                End If
                lobjDataReaderPrueba.Close()
                pstrPruebaPerfil = "Prueba"
                Exit Sub
            End If
        End If

        ' PERFILES --------------------------------------------------------------------------------------------------------------------
        If IsNumeric(pstrCodigo) Then
            Dim lobjCommandPerfiles As New OleDbCommand(sqlGetPerfilesBioquimicaByCodeLocal(), mobjCnnAccess)
            Dim lobjParamPerfiles As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
            lobjParamPerfiles.Value = CType(pstrCodigo, Integer)
            lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
            Dim lobjDataReaderPerfiles As OleDbDataReader = lobjCommandPerfiles.ExecuteReader()

            If lobjDataReaderPerfiles.Read() Then
                ' Si hay datos, pillamos la abreviación y el nombre
                ' Abreviación
                If lobjDataReaderPerfiles.IsDBNull(1) Then
                    pstrAbrv = pstrCodigo
                    pstrNombre = ""
                Else
                    pstrAbrv = pstrCodigo
                    pstrNombre = lobjDataReaderPerfiles.GetString(1)
                End If
                lobjDataReaderPerfiles.Close()
                pstrPruebaPerfil = "Perfil"
                Exit Sub
            End If
        End If

        ' MICROBIOLOGIA -----------------------------------------------------------------------------------------------------------
        Dim lobjCommandMicro As New OleDbCommand(sqlGetMicroByCodeLocalPreseleccionado(IsNumeric(pstrCodigo)), mobjCnnAccess)
        If IsNumeric(pstrCodigo) Then
            Dim lobjParamMicro As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
            lobjParamMicro.Value = CType(pstrCodigo, Integer)
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        Else
            Dim lobjParamMicro As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
            lobjParamMicro.Value = pstrCodigo
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        End If

        Dim lobjDataReaderMicro As OleDbDataReader = lobjCommandMicro.ExecuteReader()

        If lobjDataReaderMicro.Read() Then
            ' Si hay datos, pillamos la abreviación y el nombre
            ' Abreviación
            If lobjDataReaderMicro.IsDBNull(1) Then
                pstrAbrv = ""
            Else
                pstrAbrv = lobjDataReaderMicro.GetString(1)
            End If
            ' Nombre
            If lobjDataReaderMicro.IsDBNull(2) Then
                pstrNombre = ""
            Else
                pstrNombre = lobjDataReaderMicro.GetString(2)
            End If
            If IsNumeric(pstrCodigo) Then
                pstrPruebaPerfil = "Prueba"
            Else
                pstrPruebaPerfil = "Perfil"
            End If
            lobjDataReaderMicro.Close()
            Exit Sub
        End If

        pstrAbrv = ""
        pstrNombre = ""

    End Sub

    ' ************************************************************************
    ' AbrirConexion
    ' Desc: Función 
    ' NBL: 19/01/2007
    ' ************************************************************************
    Public Function AbrirConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota
            If mobjConfig.Conexion = BBDD.Local Then
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
            MessageBox.Show("No se ha podido conectar a la BBDD" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ************************************************************************
    ' CerrarConexion
    ' Desc: Función 
    ' NBL: 23/01/2007
    ' ************************************************************************
    Public Function CerrarConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota
            If mobjConfig.Conexion = BBDD.Local Then
                If Me.mobjConfig.TipoLocal = 0 Then
                    If mobjCnnAccess.State <> ConnectionState.Closed Then mobjCnnAccess.Close()
                    mobjCnnAccess.Dispose()
                    mobjCnnAccess = Nothing
                Else
                    If mobjCnnSqlServer.State <> ConnectionState.Closed Then mobjCnnSqlServer.Close()
                    mobjCnnSqlServer.Dispose()
                    mobjCnnSqlServer = Nothing
                End If
            Else
                If mobjCnnODBC.State <> ConnectionState.Closed Then mobjCnnODBC.Close()
                mobjCnnODBC.Dispose()
                mobjCnnODBC = Nothing
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("Error al cerrar la conexion a BBDD" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ************************************************************************
    ' BuscaBioquimica
    ' Desc: Rutina de busqueda de pruebas o perfiles de bioquímica
    ' NBL: 19/01/2007
    ' ************************************************************************
    Private Sub BuscaBioquimica(ByVal pstrTextoBusqueda As String)

        ' Abrimos la conexión para hacer la consulta
        If Not AbrirConexion() Then Exit Sub

        ' Aquí depende de cual sea la conexión y cual sea el tipo de busqueda tendremos que llamar a una consulta u otra
        If IsNumeric(pstrTextoBusqueda) Then
            ' En este caso la busqueda es por código -----------------------------------------------------------------------------------------------
            ' Comprobamos si la base de datos el local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getBioquimicaByCodeLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    getBioquimicaByCodeLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getBioquimicaByCodeRemota(pstrTextoBusqueda)
            End If
        Else
            ' En este caso la busqueda es por abrv o descripción -------------------------------------------------------------------------------------
            ' Comprobamos si la base de datos el local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getBioquimicaByTextLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    getBioquimicaByTextLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getBioquimicaByTextRemota(pstrTextoBusqueda)
            End If
        End If

        ' Cerramos la conexión
        CerrarConexion()

    End Sub

    ' ************************************************************************
    ' sqlGetPerfilesBioquimicaByCodeLocal
    ' Desc: Sql de bioquimica local 
    ' NBL: 24/01/2007
    ' ************************************************************************
    Private Function sqlGetPerfilesBioquimicaByCodeLocal() As String
        Dim lstrSQL As String = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA; ", _
                                                                    mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                    mobjConfigBBDD.PerfilBioquimica.TABLA)
        Return lstrSQL
    End Function

    ' ************************************************************************
    ' sqlGetMicroByCodeLocal
    ' Desc: Sql de micro local 
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMicroByCodeLocal() As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD

            If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then
                ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                ' filtrar por el código de muestra
                If mobjConfig.MuestraMicro = 0 Then
                    lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND {0} = @BUSQUEDA AND {3} = @BUSQUEDAMUESTRA ;", _
                                                        .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                        .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                        .MicroMuestra.CODIGO_PRUEBA)
                Else
                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = @BUSQUEDA ;", _
                                                        .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA)
                End If
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ;", _
                                                        .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA)
            End If

        End With

        Return lstrSQL

    End Function


    ' ************************************************************************
    ' sqlGetMicroByCodeLocalPreseleccionado
    ' Desc: Sql de micro local 
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMicroByCodeLocalPreseleccionado(ByVal pbolPrueba As Boolean) As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD
            If pbolPrueba Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = @BUSQUEDA ;", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA)
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} =  @BUSQUEDA ;", _
                                                        .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA)
            End If

        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMicroByCodeRemota
    ' Desc: Sql de micro remota
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Function sqlGetMicroByCodeRemota(ByVal pstrTextoBusqueda As String, ByVal pstrCodigoMuestra As String) As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD

            If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then
                If mobjConfig.MuestraMicro = 0 Then
                    ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                    ' filtrar por el código de muestra
                    lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND {0} = {7} AND {3} = {8}", _
                                                        .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                        .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                        .MicroMuestra.CODIGO_PRUEBA, pstrTextoBusqueda, pstrCodigoMuestra)
                Else
                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = {4}", _
                                                        .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA, pstrTextoBusqueda)
                End If
            Else
                If Me.mobjConfig.TipoConsulta = 0 Then
                    pstrTextoBusqueda += "%"
                Else
                    pstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
                End If
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} LIKE '{4}' OR {1} LIKE '{4}'", _
                                                        .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA, pstrTextoBusqueda)
            End If
        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMicroByCodeRemotaPreseleccionado
    ' Desc: Sql de micro remota
    ' NBL: 14/2/2007
    ' NBL: 07/06/2010 Nuevo parámetro de micro que lo que nos indica es que es un perfil de micro aunque el código sea numérico
    ' ************************************************************************
    Private Function sqlGetMicroByCodeRemotaPreseleccionado(ByVal pstrTextoBusqueda As String, Optional ByVal pbolPerfil As Boolean = False) As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD
            If IsNumeric(pstrTextoBusqueda) And pbolPerfil = False Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = {4}", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA, pstrTextoBusqueda)

            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = '{4}'", _
                                                        .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA, pstrTextoBusqueda)
            End If
        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMicroByTextLocal 
    ' Desc: Sql de micro local 
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMicroByTextLocal() As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD

            If clsUtil.NoCaseSensitive <> 1 Then

                If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then

                    If mobjConfig.MuestraMicro = 0 Then
                        ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                        ' filtrar por el código de muestra
                        lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND ({1} LIKE @BUSQUEDA OR {2} LIKE @BUSQUEDA)  AND {3} = @BUSQUEDAMUESTRA ;", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                            .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                            .MicroMuestra.CODIGO_PRUEBA)
                    Else
                        lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE ({1} LIKE @BUSQUEDA OR {2} LIKE @BUSQUEDA) ;", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA)
                    End If

                Else

                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ;", _
                                                .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA)
                End If

            Else

                If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then

                    If mobjConfig.MuestraMicro = 0 Then
                        ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                        ' filtrar por el código de muestra
                        lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND (UCASE({1}) LIKE @BUSQUEDA OR UCASE({2}) LIKE @BUSQUEDA)  AND UCASE({3}) = @BUSQUEDAMUESTRA ;", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                            .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                            .MicroMuestra.CODIGO_PRUEBA)
                    Else
                        lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE (UCASE({1}) LIKE @BUSQUEDA OR UCASE({2}) LIKE @BUSQUEDA) ;", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA)
                    End If

                Else

                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ;", _
                                                .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA)
                End If

            End If

        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMicroByTextRemota
    ' Desc: Sql de micro remota 
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Function sqlGetMicroByTextRemota(ByVal pstrTextoBusqueda As String, ByVal pstrCodigoMuestra As String) As String

        Dim lstrSQL As String

        ' Para construir el SQL hemos de tener en cuenta si existe una tabla que relaciones las pruebas
        ' de micro y las muestras.

        With mobjConfigBBDD

            If Me.mobjConfig.TipoConsulta = 0 Then
                pstrTextoBusqueda += "%"
            Else
                pstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
            End If

            If clsUtil.NoCaseSensitive <> 1 Then

                If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then
                    If mobjConfig.MuestraMicro = 0 Then
                        ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                        ' filtrar por el código de muestra
                        lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND ({1} LIKE '{7}' OR {2} LIKE '{7}')  AND {3} = {8}", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                            .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                            .MicroMuestra.CODIGO_PRUEBA, pstrTextoBusqueda, pstrCodigoMuestra)
                    Else
                        lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE (UCASE({1}) LIKE '{4}' OR UCASE({2}) LIKE '{4}')", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA, pstrTextoBusqueda.ToUpper)
                    End If
                Else
                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({0}) LIKE '{4}' OR UCASE({1}) LIKE '{4}'", _
                                                               .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA, pstrTextoBusqueda.ToUpper)
                End If

            Else

                If Me.txtCodigoMuestra.Text.Trim.Length > 0 Then
                    If mobjConfig.MuestraMicro = 0 Then
                        ' Construimos el string suponiendo que existe una tabla de relación micro-muestra y se tiene que 
                        ' filtrar por el código de muestra
                        lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3} FROM {4}, {5} WHERE {4}.{0} = {5}.{6} AND (UCASE({1}) LIKE '{7}' OR UCASE({2}) LIKE '{7}')  AND {3} = {8}", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, _
                                                            .MicroMuestra.CODIGO_MUESTRA, .Microbiologia.TABLA, .MicroMuestra.TABLA, _
                                                            .MicroMuestra.CODIGO_PRUEBA, pstrTextoBusqueda.ToUpper, pstrCodigoMuestra)
                    Else
                        lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE (UCASE({1}) LIKE '{4}' OR UCASE({2}) LIKE '{4}')", _
                                                            .Microbiologia.CODIGO, .Microbiologia.ABRV, .Microbiologia.NOMBRE, .Microbiologia.TABLA, pstrTextoBusqueda.ToUpper)
                    End If
                Else
                    lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({0}) LIKE '{4}' OR UCASE({1}) LIKE '{4}'", _
                                                               .PerfilMicro.CODIGO, .PerfilMicro.CODIGO, .PerfilMicro.NOMBRE, .PerfilMicro.TABLA, pstrTextoBusqueda.ToUpper)
                End If

            End If

        End With

        Return lstrSQL

    End Function


    ' ************************************************************************
    ' sqlGetMuestrasByCodeLocal
    ' Desc: Sql de muestras local 
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Function sqlGetMuestrasByCodeLocal() As String

        Dim lstrSQL As String = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE {0} = @BUSQUEDA ;", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA)
        Return lstrSQL
    End Function

    ' ************************************************************************
    ' sqlGetMuestrasByCodeRemota
    ' Desc: Sql de muestras remota
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMuestrasByCodeRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSql As String = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE {0} = {4}", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA, _
                                                                    pstrTextoBusqueda)
        Return lstrSql
    End Function

    ' ************************************************************************
    ' sqlGetMuestrasByTextRemota
    ' Desc: Sql de muestras remota
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMuestrasByTextRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSql As String

        If clsUtil.NoCaseSensitive <> 1 Then

            If mobjConfig.TipoConsulta = 0 Then
                lstrSql = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '{4}%' OR UCASE({2}) LIKE '{4}%'", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA, _
                                                                    pstrTextoBusqueda.ToUpper)
            Else
                lstrSql = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '%{4}%' OR UCASE({2}) LIKE '%{4}%'", _
                                                                            mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                            mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            End If

        Else

            If mobjConfig.TipoConsulta = 0 Then
                lstrSql = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '{4}%' OR UCASE({2}) LIKE '{4}%'", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA, _
                                                                    pstrTextoBusqueda.ToUpper)
            Else
                lstrSql = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '%{4}%' OR UCASE({2}) LIKE '%{4}%'", _
                                                                            mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                            mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            End If

        End If

        Return lstrSql

    End Function


    ' ************************************************************************
    ' sqlGetMuestrasByTextLocal
    ' Desc: Sql de muestras local por texto 
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Function sqlGetMuestrasByTextLocal() As String

        Dim lstrSQL As String = ""
        

        If clsUtil.NoCaseSensitive <> 1 Then

            lstrSQL = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE " & _
                                                                    "{1} LIKE @BUSQUEDA OR {2} LIKE @BUSQUEDA ;", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA)
        Else

            lstrSQL = String.Format("SELECT DISTINCT {0}, {1}, {2} FROM {3} WHERE " & _
                                                                    "UCASE({1}) LIKE @BUSQUEDA OR UCASE({2}) LIKE @BUSQUEDA ;", _
                                                                    mobjConfigBBDD.Muestra.CODIGO, mobjConfigBBDD.Muestra.ABRV, _
                                                                    mobjConfigBBDD.Muestra.NOMBRE, mobjConfigBBDD.Muestra.TABLA)

        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetPerfilesBioquimicaByCodeRemota
    ' Desc: Sql de bioquimica remota 
    ' NBL: 8/02/2007
    ' ************************************************************************
    Private Function sqlGetPerfilesBioquimicaByCodeRemota(ByVal pstrTextoBusqueda As String) As String
        Dim lstrSQL As String = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE {0} = {3}", _
                                                                    mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                    mobjConfigBBDD.PerfilBioquimica.TABLA, pstrTextoBusqueda)
        Return lstrSQL
    End Function

    ' ************************************************************************
    ' sqlGetPerfilesBioquimicaByTextLocal
    ' Desc: Sql de perfiles bioquimica local 
    ' NBL: 8/02/2007
    ' ************************************************************************
    Private Function sqlGetPerfilesBioquimicaByTextLocal() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE {1} LIKE @BUSQUEDA ; ", _
                                                                               mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                               mobjConfigBBDD.PerfilBioquimica.TABLA)
        Else
            lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE UCASE({1}) LIKE @BUSQUEDA ; ", _
                                                                               mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                               mobjConfigBBDD.PerfilBioquimica.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetPerfilesBioquimicaByTextRemota
    ' Desc: Sql de perfiles bioquimica remota
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Function sqlGetPerfilesBioquimicaByTextRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then

            If mobjConfig.TipoConsulta = 0 Then
                lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE UCASE({1}) LIKE '{3}%'", _
                                                                            mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                            mobjConfigBBDD.PerfilBioquimica.TABLA, pstrTextoBusqueda.ToUpper)
            Else
                lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE UCASE({1}) LIKE '%{3}%'", _
                                                                            mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                            mobjConfigBBDD.PerfilBioquimica.TABLA, pstrTextoBusqueda.ToUpper)
            End If

        Else

            If mobjConfig.TipoConsulta = 0 Then
                lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE UCASE({1}) LIKE '{3}%'", _
                                                                            mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                            mobjConfigBBDD.PerfilBioquimica.TABLA, pstrTextoBusqueda.ToUpper)
            Else
                lstrSQL = String.Format("SELECT DISTINCT {0}, {1} FROM {2} WHERE UCASE({1}) LIKE '%{3}%'", _
                                                                            mobjConfigBBDD.PerfilBioquimica.CODIGO, mobjConfigBBDD.PerfilBioquimica.NOMBRE, _
                                                                            mobjConfigBBDD.PerfilBioquimica.TABLA, pstrTextoBusqueda.ToUpper)
            End If


        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetPruebasBioquimicaByCodeLocal
    ' Desc: Sql de bioquimica local 
    ' NBL: 30/01/2007
    ' ************************************************************************
    Private Function sqlGetPruebasBioquimicaByCodeRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = {4} ", mobjConfigBBDD.Bioquimica.CODIGO, _
                                                                    mobjConfigBBDD.Bioquimica.ABRV, mobjConfigBBDD.Bioquimica.NOMBRE, _
                                                                    mobjConfigBBDD.Bioquimica.TABLA, pstrTextoBusqueda)
        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetPruebasBioquimicaByCodeLocal
    ' Desc: Sql de bioquimica local 
    ' NBL: 30/01/2007
    ' ************************************************************************
    Private Function sqlGetPruebasBioquimicaByCodeLocal() As String

        Dim lstrSQL As String = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {0} = @BUSQUEDA;", mobjConfigBBDD.Bioquimica.CODIGO, _
                                                                    mobjConfigBBDD.Bioquimica.ABRV, mobjConfigBBDD.Bioquimica.NOMBRE, _
                                                                    mobjConfigBBDD.Bioquimica.TABLA)
        Return lstrSQL
    End Function


    ' ************************************************************************
    ' sqlGetPruebasBioquimicaByTextLocal
    ' Desc: Sql de bioquimica local 
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Function sqlGetPruebasBioquimicaByTextLocal() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE {1} LIKE @BUSQUEDA OR {2} LIKE @BUSQUEDA ;", _
                                                                                mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                                mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE @BUSQUEDA OR UCASE({2}) LIKE @BUSQUEDA ;", _
                                                                                mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                                mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetPruebasBioquimicaByTextRemota
    ' Desc: Sql de bioquimica remota
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Function sqlGetPruebasBioquimicaByTextRemota(ByVal pstrTextoBusqueda As String) As String
        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then

            If mobjConfig.TipoConsulta = 0 Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '{4}%' OR UCASE({2}) LIKE '{4}%'", _
                                                                            mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                            mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '%{4}%' OR UCASE({2}) LIKE '%{4}%'", _
                                                                            mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                            mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            End If

        Else

            If mobjConfig.TipoConsulta = 0 Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '{4}%' OR UCASE({2}) LIKE '{4}%'", _
                                                                            mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                            mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3} WHERE UCASE({1}) LIKE '%{4}%' OR UCASE({2}) LIKE '%{4}%'", _
                                                                            mobjConfigBBDD.Bioquimica.CODIGO, mobjConfigBBDD.Bioquimica.ABRV, _
                                                                            mobjConfigBBDD.Bioquimica.NOMBRE, mobjConfigBBDD.Bioquimica.TABLA, _
                                                                            pstrTextoBusqueda.ToUpper)
            End If

        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' EscapeDialogoBioquimica
    ' Desc: Busqueda de pruebas o perfiles BBDD Local Access
    ' NBL: 30/01/2007
    ' ************************************************************************
    Private Sub EscapeDialogoListViewResultado(Optional ByVal pbolBorrarTexto As Boolean = True)

        ' Miramos que lv de resultado está abierto para cerrarlo y volver al textbox de búsqueda
        If Me.lvwResultadoBioquimica.Visible Then
            Me.lvwResultadoBioquimica.Visible = False
            If pbolBorrarTexto Then
                Me.txtBuscaBioquimica.ResetText()
                Me.txtBuscaBioquimica.Focus()
            End If
            Exit Sub
        End If

        If Me.lvwResultadoMuestras.Visible Then
            Me.lvwResultadoMuestras.Visible = False
            If pbolBorrarTexto Then
                Me.txtBuscaMuestraMicro.ResetText()
                Me.txtBuscaMuestraMicro.Focus()
            End If
            Exit Sub
        End If

        If Me.lvwResultadoMicro.Visible Then
            Me.lvwResultadoMicro.Visible = False
            If pbolBorrarTexto Then
                Me.txtBuscaMicro.ResetText()
                Me.txtBuscaMicro.Focus()
            End If
            Exit Sub
        End If

    End Sub

    ' ************************************************************************
    ' getBioquimicaByCodeRemota
    ' Desc: Busqueda de pruebas o perfiles BBDD Remota
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub getBioquimicaByCodeRemota(ByVal pstrTextoBusqueda As String)
        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New OdbcCommand(sqlGetPerfilesBioquimicaByCodeRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderPerfiles As OdbcDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            ' Lo metemos dentro del array
            lobjListViewItem.ImageIndex = 1
            lobjListViewItem.Tag = "Perfil"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New OdbcCommand(sqlGetPruebasBioquimicaByCodeRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderPruebas As OdbcDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------        
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 0
            lobjListViewItem.Tag = "Prueba"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMuestraByTextLocalServer
    ' Desc: Busqueda de muestras en la BBDD Local Server por texto
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub getMuestraByTextLocalServer(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New SqlCommand(sqlGetMuestrasByTextLocal(), mobjCnnSqlServer)
        Dim lobjParamMuestras As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMuestras.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMuestras.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As SqlDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()

        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el texto: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMuestraByTextLocalAccess
    ' Desc: Busqueda de muestras en la BBDD Local Access por texto
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub getMuestraByTextLocalAccess(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New OleDbCommand(sqlGetMuestrasByTextLocal(), mobjCnnAccess)
        Dim lobjParamMuestras As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMuestras.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMuestras.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As OleDbDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()

        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el texto: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMuestraByCodeLocalServer
    ' Desc: Busqueda de muestras en la BBDD Local Server
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub getMuestraByCodeLocalServer(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New SqlCommand(sqlGetMuestrasByCodeLocal(), mobjCnnSqlServer)
        Dim lobjParamMuestras As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
        lobjParamMuestras.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As SqlDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()

        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMuestraByCodeRemota
    ' Desc: Busqueda de muestras en la BBDD Remota
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub getMuestraByCodeRemota(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New OdbcCommand(sqlGetMuestrasByCodeRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderMuestras As OdbcDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()

        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMuestraByTextRemota
    ' Desc: Busqueda de muestras en la BBDD Remota
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub getMuestraByTextRemota(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New OdbcCommand(sqlGetMuestrasByTextRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderMuestras As OdbcDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()

        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el texto: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByTextLocalAccess
    ' Desc: Busqueda de pruebas de micro por texto
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Sub getMicroByTextLocalAccess(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New OleDbCommand(sqlGetMicroByTextLocal(), mobjCnnAccess)
        Dim lobjParamMicro As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMicro.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMicro.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMicro.Parameters.Add(lobjParamMicro)
        If mobjConfig.MuestraMicro = 0 And Me.txtCodigoMuestra.Text.Length > 0 Then
            Dim lobjParamMuestra As New OleDbParameter("@BUSQUEDAMUESTRA", OleDbType.Double)
            lobjParamMuestra.Value = Me.txtCodigoMuestra.Text
            lobjCommandMicro.Parameters.Add(lobjParamMuestra)
        End If
        Dim lobjDataReaderMicro As OleDbDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el texto: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByTextRemota
    ' Desc: Busqueda de pruebas de micro por texto remota
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Sub getMicroByTextRemota(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New OdbcCommand(sqlGetMicroByTextRemota(pstrTextoBusqueda, Me.txtCodigoMuestra.Text), mobjCnnODBC)
        Dim lobjDataReaderMicro As OdbcDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el texto: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByTextLocalServer
    ' Desc: Busqueda de pruebas de micro por texto server
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Sub getMicroByTextLocalServer(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New SqlCommand(sqlGetMicroByTextLocal(), mobjCnnSqlServer)
        Dim lobjParamMicro As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMicro.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMicro.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMicro.Parameters.Add(lobjParamMicro)
        If mobjConfig.MuestraMicro = 0 And Me.txtCodigoMuestra.Text.Length > 0 Then
            Dim lobjParamMuestra As New SqlParameter("@BUSQUEDAMUESTRA", SqlDbType.Float)
            lobjParamMuestra.Value = Me.txtCodigoMuestra.Text
            lobjCommandMicro.Parameters.Add(lobjParamMuestra)
        End If
        Dim lobjDataReaderMicro As SqlDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el texto: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByCodeLocalAccess
    ' Desc: Busqueda de pruebas de micro
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub getMicroByCodeLocalAccess(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New OleDbCommand(sqlGetMicroByCodeLocal(), mobjCnnAccess)

        If Me.txtCodigoMuestra.Text.Length > 0 Then
            Dim lobjParamMicro As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
            lobjParamMicro.Value = pstrTextoBusqueda
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        Else
            Dim lobjParamMicro As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
            lobjParamMicro.Value = pstrTextoBusqueda
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        End If

        If Me.txtCodigoMuestra.Text.Length > 0 And mobjConfig.MuestraMicro = 0 Then
            Dim lobjParamMuestra As New OleDbParameter("@BUSQUEDAMUESTRA", OleDbType.Double)
            lobjParamMuestra.Value = Me.txtCodigoMuestra.Text
            lobjCommandMicro.Parameters.Add(lobjParamMuestra)
        End If

        Dim lobjDataReaderMicro As OleDbDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el código: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByCodeRemota
    ' Desc: Busqueda de pruebas de micro remota
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Sub getMicroByCodeRemota(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New OdbcCommand(sqlGetMicroByCodeRemota(pstrTextoBusqueda, Me.txtCodigoMuestra.Text), mobjCnnODBC)
        Dim lobjDataReaderMicro As OdbcDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el código: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getMicroByCodeLocalServer
    ' Desc: Busqueda de pruebas de micro server
    ' NBL: 14/2/2007
    ' ************************************************************************
    Private Sub getMicroByCodeLocalServer(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMicro As New SqlCommand(sqlGetMicroByCodeLocal(), mobjCnnSqlServer)

        If Me.txtCodigoMuestra.Text.Length > 0 Then
            Dim lobjParamMicro As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
            lobjParamMicro.Value = pstrTextoBusqueda
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        Else
            Dim lobjParamMicro As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
            lobjParamMicro.Value = pstrTextoBusqueda
            lobjCommandMicro.Parameters.Add(lobjParamMicro)
        End If

        If mobjConfig.MuestraMicro = 0 And Me.txtCodigoMuestra.Text.Length > 0 Then
            Dim lobjParamMuestra As New SqlParameter("@BUSQUEDAMUESTRA", SqlDbType.Float)
            lobjParamMuestra.Value = Me.txtCodigoMuestra.Text
            lobjCommandMicro.Parameters.Add(lobjParamMuestra)
        End If
        Dim lobjDataReaderMicro As SqlDataReader = lobjCommandMicro.ExecuteReader()

        Dim larlstMicro As New ArrayList()
        While lobjDataReaderMicro.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMicro.GetValue(0), String))
            ' Abreviación de micro
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMicro.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMicro.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            If Me.txtCodigoMuestra.Text.Length > 0 Then
                lobjListViewItem.ImageIndex = 0
                lobjListViewItem.Tag = "Prueba"
            Else
                lobjListViewItem.ImageIndex = 1
                lobjListViewItem.Tag = "Perfil"
            End If
            larlstMicro.Add(lobjListViewItem)
        End While

        lobjDataReaderMicro.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de micro
        If larlstMicro.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMicro(larlstMicro)
        Else
            MessageBox.Show("No hay ninguna prueba de micro con el código: " + pstrTextoBusqueda, Me.Text, _
                                         MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' RellenaListViewMicro
    ' Desc: Rutina que rellena el listview de micro
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub RellenaListViewMicro(ByVal parlstMicro As ArrayList)

        With Me.lvwResultadoMicro

            .Items.Clear()
            .BeginUpdate()
            .SuspendLayout()
            .Items.AddRange(parlstMicro.ToArray(GetType(ListViewItem)))
            .EndUpdate()
            .ResumeLayout()

            .Visible = True
            .BringToFront()

            .Focus()
            .Items(0).Selected = True
            .Items(0).Focused = True

        End With

    End Sub

    ' ************************************************************************
    ' getMuestraByCodeLocalAccess
    ' Desc: Busqueda de muestras en la BBDD Local Access
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub getMuestraByCodeLocalAccess(ByVal pstrTextoBusqueda As String)

        Dim lobjCommandMuestras As New OleDbCommand(sqlGetMuestrasByCodeLocal(), mobjCnnAccess)
        Dim lobjParamMuestras As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
        lobjParamMuestras.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandMuestras.Parameters.Add(lobjParamMuestras)
        Dim lobjDataReaderMuestras As OleDbDataReader = lobjCommandMuestras.ExecuteReader()

        Dim larlstMuestras As New ArrayList()
        While lobjDataReaderMuestras.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMuestras.GetValue(0), String))
            ' Abreviación de la muestra
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(1), String))
            ' Descripción
            If Not lobjDataReaderMuestras.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMuestras.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 2
            larlstMuestras.Add(lobjListViewItem)
        End While

        lobjDataReaderMuestras.Close()

        ' Si hay datos lo metemos dentro del lvw de resultado de muestras
        If larlstMuestras.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewMuestras(larlstMuestras)
        Else
            MessageBox.Show("No hay ninguna muestra con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' RellenaListViewMuestras
    ' Desc: Rutina que rellena el listview de 
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub RellenaListViewMuestras(ByVal parlstMuestras As ArrayList)

        With Me.lvwResultadoMuestras

            .Items.Clear()
            .BeginUpdate()
            .SuspendLayout()
            .Items.AddRange(parlstMuestras.ToArray(GetType(ListViewItem)))
            .EndUpdate()
            .ResumeLayout()

            .Visible = True
            .BringToFront()

            .Focus()
            .Items(0).Selected = True
            .Items(0).Focused = True

        End With

    End Sub

    ' ************************************************************************
    ' getBioquimicaByCodeLocalAccess
    ' Desc: Busqueda de pruebas o perfiles BBDD Local Access
    ' NBL: 24/01/2007
    ' ************************************************************************
    Private Sub getBioquimicaByCodeLocalAccess(ByVal pstrTextoBusqueda As String)

        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New OleDbCommand(sqlGetPerfilesBioquimicaByCodeLocal(), mobjCnnAccess)
        Dim lobjParamPerfiles As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
        lobjParamPerfiles.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
        Dim lobjDataReaderPerfiles As OleDbDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList()

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            ' Importante, ponemos que es un perfil
            lobjListViewItem.Tag = "Perfil"
            lobjListViewItem.ImageIndex = 1

            ' Lo metemos dentro del array
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New OleDbCommand(sqlGetPruebasBioquimicaByCodeLocal(), mobjCnnAccess)
        Dim lobjParamPruebas As New OleDbParameter("@BUSQUEDA", OleDbType.Double)
        lobjParamPruebas.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandPruebas.Parameters.Add(lobjParamPruebas)
        Dim lobjDataReaderPruebas As OleDbDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------        
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.Tag = "Prueba"
            lobjListViewItem.ImageIndex = 0
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getBioquimicaByCodeLocalServer
    ' Desc: Busqueda de pruebas o perfiles BBDD Local Server
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub getBioquimicaByCodeLocalServer(ByVal pstrTextoBusqueda As String)

        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New SqlCommand(sqlGetPerfilesBioquimicaByCodeLocal(), mobjCnnSqlServer)
        Dim lobjParamPerfiles As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
        lobjParamPerfiles.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
        Dim lobjDataReaderPerfiles As SqlDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 1
            lobjListViewItem.Tag = "Perfil"

            ' Lo metemos dentro del array
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New SqlCommand(sqlGetPruebasBioquimicaByCodeLocal(), mobjCnnSqlServer)
        Dim lobjParamPruebas As New SqlParameter("@BUSQUEDA", SqlDbType.Float)
        lobjParamPruebas.Value = CType(pstrTextoBusqueda, Integer)
        lobjCommandPruebas.Parameters.Add(lobjParamPruebas)
        Dim lobjDataReaderPruebas As SqlDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------        
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 0
            lobjListViewItem.Tag = "Prueba"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' AddResultadoMicroTotal
    ' Desc: Aceptar resultado prueba microbiología
    ' NBL: 12/02/2007
    ' ************************************************************************
    Private Sub AddResultadoMicroTotal()

        ' Miramos que haya algo seleccionado
        If Not Me.lvwResultadoMicro.SelectedItems(0) Is Nothing Then

            Dim lobjSelected As ListViewItem = Me.lvwResultadoMicro.SelectedItems(0)

            Dim lstrCodigoPrueba As String = lobjSelected.Text
            Dim lstrAbrvPrueba As String = lobjSelected.SubItems(1).Text
            Dim lstrNombrePrueba As String = lobjSelected.SubItems(2).Text

            If Not ExistePruebaMuestra(lstrCodigoPrueba, Me.txtCodigoMuestra.Text.Trim) Then
                Me.dgvPruebasTotal.Rows.Add(lstrCodigoPrueba, lstrAbrvPrueba, lstrNombrePrueba, Me.txtCodigoMuestra.Text.Trim, _
                                                            Me.txtAbrvMuestra.Text.Trim, Me.txtNombreMuestra.Text.Trim, "M", lobjSelected.Tag)
                Me.dgvPruebasTotal.FirstDisplayedScrollingRowIndex = Me.dgvPruebasTotal.Rows.Count - 1
            End If

            Me.txtBuscaMicro.ResetText()
            ' Lo hacemos invisible      
            Me.lvwResultadoMicro.Visible = False
            Me.txtBuscaMicro.Focus()
        End If

        Me.dgvPruebasTotal.ClearSelection()

    End Sub

    ' ************************************************************************
    ' AddResultadoBioquimicaTotal
    ' Desc: Aceptar resultado prueba bioquímica
    ' NBL: 7/02/2007
    ' ************************************************************************
    Private Sub AddResultadoBioquimicaTotal()

        ' Miramos si hay algo seleccionado
        If Me.lvwResultadoBioquimica.SelectedItems.Count > 0 Then

            Dim lobjSelected As ListViewItem = Me.lvwResultadoBioquimica.SelectedItems(0)

            Dim lstrCodigoPrueba As String = lobjSelected.Text
            Dim lstrAbrvPrueba As String = lobjSelected.SubItems(1).Text
            Dim lstrNombrePrueba As String = lobjSelected.SubItems(2).Text

            If Not ExistePruebaMuestra(lstrCodigoPrueba, "") Then
                Me.dgvPruebasTotal.Rows.Add(lstrCodigoPrueba, lstrAbrvPrueba, lstrNombrePrueba, "", "", "", "B", lobjSelected.Tag)
                ' Mostramos la última prueba seleccionada
                Me.dgvPruebasTotal.FirstDisplayedScrollingRowIndex = Me.dgvPruebasTotal.Rows.Count - 1
                Me.dgvPruebasTotal.ClearSelection()
            End If

            ' Borramos el text box de busqueda bioquímica
            Me.txtBuscaBioquimica.ResetText()
            Me.lvwResultadoBioquimica.Visible = False
            Me.txtBuscaBioquimica.Focus()

        End If

        Me.dgvPruebasTotal.ClearSelection()

    End Sub

    ' ************************************************************************
    ' ExistePruebaMuestra
    ' Desc: Comprobamos si la combinación de prueba y muestra existe en el grid total
    ' NBL: 7/02/2007
    ' ************************************************************************
    Private Function ExistePruebaMuestra(ByVal pstrCodigoPrueba As String, ByVal pstrCodigoMuestra As String) As Boolean

        If Me.dgvPruebasTotal.Rows.Count > 0 Then
            For lintContador As Integer = 0 To Me.dgvPruebasTotal.Rows.Count - 1
                Dim FilaDataGrid As DataGridViewRow = Me.dgvPruebasTotal.Rows(lintContador)
                If pstrCodigoPrueba = CType(FilaDataGrid.Cells("CodigoT").Value, String) And _
                        pstrCodigoMuestra = CType(FilaDataGrid.Cells("CodigoM").Value, String) Then Return True
            Next
        End If

        Return False

    End Function

    ' ************************************************************************
    ' AceptarResultadoMuestra
    ' Desc: Aceptar resultado muestra
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub AceptarResultadoMuestra()

        ' Miramos que haya algo seleccionado
        If Not Me.lvwResultadoMuestras.SelectedItems(0) Is Nothing Then
            With Me.lvwResultadoMuestras.SelectedItems(0)
                Me.txtCodigoMuestra.Text = .Text
                Me.txtAbrvMuestra.Text = .SubItems(1).Text
                Me.txtNombreMuestra.Text = .SubItems(2).Text
                ' Importante, para la tecla rápida de aceptar
                Me.txtBuscaMuestraMicro.Tag = "1"
            End With
            ' Lo hacemos invisible
            Me.lvwResultadoMuestras.Visible = False
            ' Como solo se puede capturar una muestra por cada prueba de micro,
            ' lo que hacemos es pasar el foco al busca micro
            Me.txtBuscaMicro.Focus()
            ' Reseteamos el texto del campo de busca muestra
            Me.txtBuscaMuestraMicro.ResetText()
        End If

    End Sub

    ' ************************************************************************
    ' AceptarResultadoBioquimica
    ' Desc: Aceptar resultado prueba bioquímica
    ' NBL: 30/01/2007
    ' ************************************************************************
    'Private Sub AceptarResultadoBioquimica()

    '    ' Miramos que haya algo seleccionado
    '    If Not Me.lvwResultadoBioquimica.SelectedItems(0) Is Nothing Then
    '        With Me.lvwResultadoBioquimica.SelectedItems(0)
    '            Me.txtCodigoBioquimica.Text = .Text
    '            Me.txtAbrvBioquimica.Text = .SubItems(1).Text
    '            Me.txtNombreBioquimica.Text = .SubItems(2).Text
    '            ' Importante, para la tecla rápida de aceptar
    '            Me.txtBuscaBioquimica.Tag = "1"
    '        End With
    '        ' Lo hacemos invisible
    '        Me.lvwResultadoBioquimica.Visible = False
    '        Me.txtBuscaBioquimica.Focus()
    '    End If

    'End Sub


    ' ************************************************************************
    ' RellenaListViewPruebasBiquimica
    ' Desc: Rutina que rellena el listview de 
    ' NBL: 30/01/2007
    ' ************************************************************************
    Private Sub RellenaListViewPruebasBiquimica(ByVal parlstPruebas As ArrayList)

        With Me.lvwResultadoBioquimica

            .Items.Clear()
            .BeginUpdate()
            .SuspendLayout()
            .Items.AddRange(parlstPruebas.ToArray(GetType(ListViewItem)))
            .EndUpdate()
            .ResumeLayout()

            ' Ponemos el visible el listview y lo situamos delante de todo
            .Visible = True
            .BringToFront()

            .Focus()
            .Items(0).Selected = True
            .Items(0).Focused = True

        End With

    End Sub

    Private Sub DialogoPruebas_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

    End Sub

    Private Sub DialogoPruebas_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If Not Me.mbolAceptandoDialogo Then

            ' Si no hay ninguno es que hay que cerrar el diálogo
            If MessageBox.Show("Cancelar la selección de pruebas", Me.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question, _
                MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                e.Cancel = True
            Else
                Me.Resultado = Me.pstrSeleccionEntrada
                Me.DialogResult = Windows.Forms.DialogResult.OK
            End If

        End If

        ' Cargamos las dimensiones del diálogo
        If Me.Top > 0 Then
            My.Settings.DialogoPruebasLocation = Me.Location
        End If

        If Me.WindowState <> FormWindowState.Maximized Then My.Settings.DialogoPruebasSize = Me.Size
        My.Settings.DialogoPruebasState = Me.WindowState

        ' Cargamos las columnas del diálogo de pruebas
        My.Settings.DialogoPruebasCodigoWidth = Me.lvwResultadoBioquimica.Columns(0).Width
        My.Settings.DialogoPruebasAbrvWidth = Me.lvwResultadoBioquimica.Columns(1).Width
        My.Settings.DialogoPruebasNombreWidth = Me.lvwResultadoBioquimica.Columns(2).Width

        My.Settings.DialogoPruebasCodigoTWidth = Me.dgvPruebasTotal.Columns(0).Width
        My.Settings.DialogoPruebasAbrvTWidth = Me.dgvPruebasTotal.Columns(1).Width
        My.Settings.DialogoPruebasNombreTWidth = Me.dgvPruebasTotal.Columns(2).Width
        My.Settings.DialogoPruebasMuestraTWidth = Me.dgvPruebasTotal.Columns(3).Width
        My.Settings.DialogoPruebasMuestra2TWidth = Me.dgvPruebasTotal.Columns(4).Width
        My.Settings.DialogoPruebasMuestra3TWidth = Me.dgvPruebasTotal.Columns(5).Width

        My.Settings.DialogoPruebasCodigoMuestraWidth = Me.lvwResultadoMuestras.Columns(0).Width
        My.Settings.DialogoPruebasAbrvMuestraWidth = Me.lvwResultadoMuestras.Columns(1).Width
        My.Settings.DialogoPruebasNombreMuestraWidth = Me.lvwResultadoMuestras.Columns(2).Width

        My.Settings.DialogoPruebasCodigoMicroWidth = Me.lvwResultadoMicro.Columns(0).Width
        My.Settings.DialogoPruebasAbrvMicroWidth = Me.lvwResultadoMicro.Columns(1).Width
        My.Settings.DialogoPruebasNombreMicroWidth = Me.lvwResultadoMicro.Columns(2).Width

        My.Settings.Save()

    End Sub

    Private Sub DialogoPruebas_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Escape Then EscapeDialogoListViewResultado()

        If e.KeyCode = Keys.F12 Then AceptarResultadoFinal()

    End Sub

    Private Sub DialogoPruebas_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.dgvPruebasTotal.ClearSelection()

        ' Cargamos las dimensiones del diálogo
        If Not My.Settings.DialogoPruebasLocation.IsEmpty Then
            Me.Location = My.Settings.DialogoPruebasLocation
            If Me.Top < 0 Then
                Me.Top = 0
                Me.Left = 0
            End If
        End If

        Me.Size = My.Settings.DialogoPruebasSize

        If My.Settings.DialogoPruebasState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Maximized
        Else
            Me.WindowState = FormWindowState.Normal
        End If

        If Me.Width < 300 Then Me.Width = 300
        If Me.Height < 300 Then Me.Height = 300

        ' Cargamos las columnas del diálogo de pruebas
        Me.lvwResultadoBioquimica.Columns(0).Width = IIf(My.Settings.DialogoPruebasCodigoWidth = 0, 50, My.Settings.DialogoPruebasCodigoWidth)
        Me.lvwResultadoBioquimica.Columns(1).Width = IIf(My.Settings.DialogoPruebasAbrvWidth = 0, 50, My.Settings.DialogoPruebasAbrvWidth)
        Me.lvwResultadoBioquimica.Columns(2).Width = IIf(My.Settings.DialogoPruebasNombreWidth = 0, 50, My.Settings.DialogoPruebasNombreWidth)

        Me.dgvPruebasTotal.Columns(0).Width = IIf(My.Settings.DialogoPruebasCodigoTWidth = 0, 50, My.Settings.DialogoPruebasCodigoTWidth)
        Me.dgvPruebasTotal.Columns(1).Width = IIf(My.Settings.DialogoPruebasAbrvTWidth = 0, 50, My.Settings.DialogoPruebasAbrvTWidth)
        Me.dgvPruebasTotal.Columns(2).Width = IIf(My.Settings.DialogoPruebasNombreTWidth = 0, 50, My.Settings.DialogoPruebasNombreTWidth)
        Me.dgvPruebasTotal.Columns(3).Width = IIf(My.Settings.DialogoPruebasMuestraTWidth = 0, 50, My.Settings.DialogoPruebasMuestraTWidth)
        Me.dgvPruebasTotal.Columns(4).Width = IIf(My.Settings.DialogoPruebasMuestra2TWidth = 0, 50, My.Settings.DialogoPruebasMuestra2TWidth)
        Me.dgvPruebasTotal.Columns(5).Width = IIf(My.Settings.DialogoPruebasMuestra3TWidth = 0, 50, My.Settings.DialogoPruebasMuestra3TWidth)

        Me.lvwResultadoMuestras.Columns(0).Width = IIf(My.Settings.DialogoPruebasCodigoMuestraWidth = 0, 50, My.Settings.DialogoPruebasCodigoMuestraWidth)
        Me.lvwResultadoMuestras.Columns(1).Width = IIf(My.Settings.DialogoPruebasAbrvMuestraWidth = 0, 50, My.Settings.DialogoPruebasAbrvMuestraWidth)
        Me.lvwResultadoMuestras.Columns(2).Width = IIf(My.Settings.DialogoPruebasNombreMuestraWidth = 0, 50, My.Settings.DialogoPruebasNombreMuestraWidth)

        Me.lvwResultadoMicro.Columns(0).Width = IIf(My.Settings.DialogoPruebasCodigoMicroWidth = 0, 50, My.Settings.DialogoPruebasCodigoMicroWidth)
        Me.lvwResultadoMicro.Columns(1).Width = IIf(My.Settings.DialogoPruebasAbrvMicroWidth = 0, 50, My.Settings.DialogoPruebasAbrvMicroWidth)
        Me.lvwResultadoMicro.Columns(2).Width = IIf(My.Settings.DialogoPruebasNombreMicroWidth = 0, 50, My.Settings.DialogoPruebasNombreMicroWidth)

    End Sub

    Private Sub txtBuscaBioquimica_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBuscaBioquimica.GotFocus

        EscapeDialogoListViewResultado(False)

    End Sub

    Private Sub txtBuscaBioquimica_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscaBioquimica.KeyDown

        If e.KeyCode = Keys.Enter And Me.txtBuscaBioquimica.Text.Trim.Length > 0 Then
            ' Pulsan enter y tenemos datos en el textbox de busqueda
            Try
                BuscaBioquimica(Me.txtBuscaBioquimica.Text.Trim.ToUpper)
            Catch ex As Exception
                MessageBox.Show("Error en la busqueda de pruebas/perfiles de bioquímica. " + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub

    Private Sub lvwResultadoBioquimica_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvwResultadoBioquimica.KeyDown

        If e.KeyCode = Keys.Enter Then
            'AceptarResultadoBioquimica()
            AddResultadoBioquimicaTotal()
        End If

    End Sub



    Private Sub txtBuscaMuestraMicro_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBuscaMuestraMicro.GotFocus

        If Me.txtCodigoMuestra.Text.Trim.Length = 0 Then EscapeDialogoListViewResultado(False)

    End Sub

    Private Sub txtBuscaMuestraMicro_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscaMuestraMicro.KeyDown

        If e.KeyCode = Keys.Enter And Me.txtBuscaMuestraMicro.Text.Trim.Length > 0 Then
            ' Pulsamos enter y tenemos datos en el textbox de busqueda
            Try
                BuscaMuestra(Me.txtBuscaMuestraMicro.Text.Trim.ToUpper)
            Catch ex As Exception
                MessageBox.Show("Error en la busqueda de muestras de microbiología" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        Else
            resetTextBoxesMuestra()
        End If

    End Sub

    ' ************************************************************************
    ' BuscaMicro
    ' Desc: Rutina de busqueda de pruebas de microbiología
    ' NBL: 12/2/2007
    ' ************************************************************************
    Private Sub BuscaMicro(ByVal pstrTextoBusqueda As String)

        ' Abrimos la conexión para hacer la consulta
        If Not AbrirConexion() Then Exit Sub

        ' Aquí depende de cual sea la conexión y cual sea el tipo de busqueda tendremos que llamar  a una consulta u otra
        If IsNumeric(pstrTextoBusqueda) Then
            ' Busqueda por código --------------------------------------------------------------------------
            ' Comprobamos si la base de datos es local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getMicroByCodeLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL Server
                    getMicroByCodeRemota(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getMicroByCodeRemota(pstrTextoBusqueda)
            End If
        Else
            ' En este caso la busqueda es por abrv o descripción -------------------------------------------
            ' Comprobamos si la base de datos es local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getMicroByTextLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    getMicroByTextLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getMicroByTextRemota(pstrTextoBusqueda)
            End If

        End If

        ' Cerramos la conexión
        CerrarConexion()

    End Sub

    ' ************************************************************************
    ' BuscaMuestra
    ' Desc: Rutina de busqueda de muestras
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub BuscaMuestra(ByVal pstrTextoBusqueda As String)

        ' Abrimos la conexión para hacer la consulta
        If Not AbrirConexion() Then Exit Sub

        ' Aquí depende de cual sea la conexión y cual sea el tipo de busqueda tendremos que llamar  a una consulta u otra
        If IsNumeric(pstrTextoBusqueda) Then
            ' Busqueda por código --------------------------------------------------------------------------
            ' Comprobamos si la base de datos es local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getMuestraByCodeLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL Server
                    getMuestraByCodeLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getMuestraByCodeRemota(pstrTextoBusqueda)
            End If
        Else
            ' En este caso la busqueda es por abrv o descripción -------------------------------------------
            ' Comprobamos si la base de datos es local o remota
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    getMuestraByTextLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    getMuestraByTextLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                getMuestraByTextRemota(pstrTextoBusqueda)
            End If

        End If

        ' Cerramos la conexión
        CerrarConexion()

    End Sub

    ' ************************************************************************
    ' resetTextBoxesMuestra
    ' Desc: Reseteamos los textboxes que muestran las muestras
    ' NBL: 9/2/2007
    ' ************************************************************************
    Private Sub resetTextBoxesMuestra()

        Me.txtCodigoMuestra.Text = ""
        Me.txtAbrvMuestra.Text = ""
        Me.txtNombreMuestra.Text = ""

    End Sub

    ' ************************************************************************
    ' getBioquimicaByTextRemota
    ' Desc: Busqueda de pruebas/perfiles por texto remota
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Sub getBioquimicaByTextRemota(ByVal pstrTextoBusqueda As String)

        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New OdbcCommand(sqlGetPerfilesBioquimicaByTextRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderPerfiles As OdbcDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 1
            lobjListViewItem.Tag = "Perfil"
            ' Lo metemos dentro del array
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New OdbcCommand(sqlGetPruebasBioquimicaByTextRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderPruebas As OdbcDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 0
            lobjListViewItem.Tag = "Prueba"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getBioquimicaByTextLocalServer
    ' Desc: Busqueda de pruebas/perfiles por texto local server
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Sub getBioquimicaByTextLocalServer(ByVal pstrTextoBusqueda As String)

        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New SqlCommand(sqlGetPerfilesBioquimicaByTextLocal(), mobjCnnSqlServer)
        Dim lobjParamPerfiles As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If mobjConfig.TipoConsulta = 0 Then
            lobjParamPerfiles.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamPerfiles.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
        Dim lobjDataReaderPerfiles As SqlDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 1
            lobjListViewItem.Tag = "Perfil"
            ' Lo metemos dentro del array
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New SqlCommand(sqlGetPruebasBioquimicaByTextLocal(), mobjCnnSqlServer)
        Dim lobjParamPruebas As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        lobjParamPruebas.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        lobjCommandPruebas.Parameters.Add(lobjParamPruebas)
        Dim lobjDataReaderPruebas As SqlDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 0
            lobjListViewItem.Tag = "Prueba"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************
    ' getBioquimicaByTextLocalAccess
    ' Desc: Busqueda de pruebas/perfiles por texto local access
    ' NBL: 8/2/2007
    ' ************************************************************************
    Private Sub getBioquimicaByTextLocalAccess(ByVal pstrTextoBusqueda As String)

        ' Construimos la instrucción SQL para los perfiles
        Dim lobjCommandPerfiles As New OleDbCommand(sqlGetPerfilesBioquimicaByTextLocal(), mobjCnnAccess)
        Dim lobjParamPerfiles As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If mobjConfig.TipoConsulta = 0 Then
            lobjParamPerfiles.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamPerfiles.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If

        lobjCommandPerfiles.Parameters.Add(lobjParamPerfiles)
        Dim lobjDataReaderPerfiles As OleDbDataReader = lobjCommandPerfiles.ExecuteReader()

        Dim larlstPruebas As New ArrayList

        ' 1.- Primero pillamos los perfiles de bioquímica ----------------------------------------------------------------------------
        While lobjDataReaderPerfiles.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Como el perfil no tiene abreviación lo que hacemos es volver a poner el código
            lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(0), String))
            ' Ponemos la descripción del perfil
            If Not lobjDataReaderPerfiles.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPerfiles.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 1
            lobjListViewItem.Tag = "Perfil"
            ' Lo metemos dentro del array
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPerfiles.Close()

        ' Construimos la instrucción SQL para las pruebas vulgaris
        Dim lobjCommandPruebas As New OleDbCommand(sqlGetPruebasBioquimicaByTextLocal(), mobjCnnAccess)
        Dim lobjParamPruebas As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        lobjParamPruebas.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        lobjCommandPruebas.Parameters.Add(lobjParamPruebas)
        Dim lobjDataReaderPruebas As OleDbDataReader = lobjCommandPruebas.ExecuteReader()

        ' 2.- Seguidamente pillamos las pruebas vulgaris ----------------------------------------------------------------------------
        While lobjDataReaderPruebas.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderPruebas.GetValue(0), String))
            ' Abreviación
            If Not lobjDataReaderPruebas.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            ' Descripción
            If Not lobjDataReaderPruebas.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderPruebas.GetValue(2), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text
            lobjListViewItem.ImageIndex = 0
            lobjListViewItem.Tag = "Prueba"
            larlstPruebas.Add(lobjListViewItem)
        End While

        lobjDataReaderPruebas.Close()

        If larlstPruebas.Count > 0 Then
            ' Ponemos los datos en el listview
            RellenaListViewPruebasBiquimica(larlstPruebas)
        Else
            MessageBox.Show("No hay ninguna prueba/perfil con el código: " + pstrTextoBusqueda, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If


    End Sub

    Private Sub txtBuscaMicro_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtBuscaMicro.GotFocus

        EscapeDialogoListViewResultado(False)

    End Sub

    Private Sub txtBuscaMicro_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscaMicro.KeyDown

        If e.KeyCode = Keys.Enter And Me.txtBuscaMicro.Text.Trim.Length > 0 Then
            ' Pulsamos enter y tenemos datos en el textbox de busqueda
            Try
                BuscaMicro(Me.txtBuscaMicro.Text.Trim.ToUpper)
            Catch ex As Exception
                MessageBox.Show("Error en la busqueda de pruebas de microbiología" + vbCrLf + ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        End If

    End Sub

    Private Sub lvwResultadoMuestras_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvwResultadoMuestras.KeyDown

        If e.KeyCode = Keys.Enter Then
            AceptarResultadoMuestra()
        End If

    End Sub

    Private Sub txtBuscaMicro_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscaMicro.TextChanged

    End Sub

    Private Sub lvwResultadoMicro_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvwResultadoMicro.KeyDown
        If e.KeyCode = Keys.Enter Then
            AddResultadoMicroTotal()
        End If
    End Sub


    ' ************************************************************************
    ' AceptarResultadoFinal
    ' Desc: Aceptar resultado final
    ' NBL: 14/02/2007
    ' ************************************************************************
    Private Sub AceptarResultadoFinal()

        If Me.dgvPruebasTotal.Rows.Count > 0 Then
            Dim lstrBResultado As New StringBuilder
            ' Hacemos el bucle por todas las rows que haya
            For Each lobjRow As DataGridViewRow In Me.dgvPruebasTotal.Rows
                lstrBResultado.Append(",")
                lstrBResultado.Append(CType(lobjRow.Cells("CodigoT").Value, String))
                lstrBResultado.Append("^")
                lstrBResultado.Append(CType(lobjRow.Cells("TipoPrueba").Value, String))
                lstrBResultado.Append("|")
                lstrBResultado.Append(CType(lobjRow.Cells("CodigoM").Value, String).Trim)
                lstrBResultado.Append("|")
            Next
            Me.Resultado = lstrBResultado.ToString()
        Else
            Me.Resultado = ""
        End If

        Me.mbolAceptandoDialogo = True
        Me.DialogResult = Windows.Forms.DialogResult.OK

        Me.Close()

    End Sub

    Private Sub dgvPruebasTotal_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles dgvPruebasTotal.CellFormatting

        Dim lobjDGVR As DataGridViewRow = Me.dgvPruebasTotal.Rows(e.RowIndex)
        Dim lobjColor As System.Drawing.Color

        Dim lstrPruebaPerfil As String = CType(Me.dgvPruebasTotal.Item(7, e.RowIndex).Value, String)
        Dim lstrMuestra As String = CType(Me.dgvPruebasTotal.Item(4, e.RowIndex).Value, String)
        Dim lstrTipoPrueba As String = CType(Me.dgvPruebasTotal.Item(6, e.RowIndex).Value, String)

        If lstrTipoPrueba = "M" And lstrPruebaPerfil = "Prueba" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
            lobjColor = Drawing.Color.PaleGreen
        ElseIf lstrTipoPrueba = "M" And lstrPruebaPerfil = "Perfil" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
            lobjColor = Drawing.Color.LimeGreen
        ElseIf lstrTipoPrueba = "B" And lstrPruebaPerfil = "Prueba" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
            lobjColor = Drawing.Color.MistyRose
        ElseIf lstrTipoPrueba = "B" And lstrPruebaPerfil = "Perfil" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
            lobjColor = Drawing.Color.LightCoral
        ElseIf lstrMuestra.Length > 0 And (e.ColumnIndex = 3 Or e.ColumnIndex = 4 Or e.ColumnIndex = 5) Then
            lobjColor = Drawing.Color.LightSkyBlue
        Else
            lobjColor = Drawing.Color.White
        End If

        'If lstrPruebaPerfil = "Prueba" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
        '    lobjColor = Drawing.Color.PaleGreen
        'ElseIf lstrPruebaPerfil = "Perfil" And (e.ColumnIndex = 0 Or e.ColumnIndex = 1 Or e.ColumnIndex = 2) Then
        '    lobjColor = Drawing.Color.LightCoral
        'ElseIf lstrMuestra.Length > 0 And (e.ColumnIndex = 3 Or e.ColumnIndex = 4 Or e.ColumnIndex = 5) Then
        '    lobjColor = Drawing.Color.LightSkyBlue
        'Else
        '    lobjColor = Drawing.Color.White
        'End If

        e.CellStyle.BackColor = lobjColor

    End Sub

    Private Sub dgvPruebasTotal_RowsRemoved(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsRemovedEventArgs) Handles dgvPruebasTotal.RowsRemoved
        Me.dgvPruebasTotal.Refresh()
    End Sub

    Private Sub dgvPruebasTotal_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPruebasTotal.SelectionChanged

    End Sub

    Private Sub dgvPruebasTotal_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvPruebasTotal.CellContentClick

    End Sub

    Private Sub dgvPruebasTotal_Sorted(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvPruebasTotal.Sorted
        Me.dgvPruebasTotal.Refresh()
    End Sub

    Private Sub btnBorrarTodo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBorrarTodo.Click

        If MessageBox.Show("¿Desea borrar todas las pruebas seleccionadas?", Me.Text, MessageBoxButtons.YesNo, _
                                            MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            ' Borramos todo lo que haya seleccionado 
            Me.dgvPruebasTotal.Rows.Clear()
        End If
    End Sub

    Private Sub btnAceptar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAceptar.Click
        AceptarResultadoFinal()
    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click
        Me.Close()
    End Sub

    Private Sub txtBuscaBioquimica_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscaBioquimica.TextChanged

    End Sub

    Private Sub lvwResultadoBioquimica_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvwResultadoBioquimica.MouseDoubleClick

        AddResultadoBioquimicaTotal()

    End Sub

    Private Sub lvwResultadoMuestras_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvwResultadoMuestras.MouseDoubleClick

        AceptarResultadoMuestra()

    End Sub

    Private Sub lvwResultadoMicro_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvwResultadoMicro.MouseDoubleClick
        AddResultadoMicroTotal()
    End Sub

    Private Sub txtBuscaMuestraMicro_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscaMuestraMicro.TextChanged

    End Sub
End Class