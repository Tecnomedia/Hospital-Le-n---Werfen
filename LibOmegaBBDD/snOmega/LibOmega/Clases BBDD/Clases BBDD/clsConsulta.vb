Imports System.Data.OleDb
Imports System.Data.Odbc
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Text

Public Class clsConsulta

    Public Declare Function MSDPE Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" ( _
    ByVal lpPath As String) As Long

    ' Instancias de configuración
    Private mobjConfig As clsConfig
    Private mobjConfigBBDD As clsConfigBBDD

    ' Ponemos las tres tipos de conexiones que puede haber en este formulario
    Private mobjCnnAccess As OleDbConnection
    Private mobjCnnSqlServer As SqlConnection
    Private mobjCnnODBC As OdbcConnection
    ' NBL 26/6/2009 Para la carga de pruebas con el FlexibarNET
    Private mobjPruebas As DialogoPruebas

    ' ********************************************************************************
    ' IniciarCargaPruebas
    ' Desc: Rutina que inicializa la variable para la carga de pruebas
    ' NBL: 25/6/2009
    ' ********************************************************************************
    Public Sub IniciarCargaPruebas()

        mobjPruebas = New DialogoPruebas("", "")

        AbrirConexion()

    End Sub

    ' ********************************************************************************
    ' CerrarCargaPruebas
    ' Desc: Rutina que termina la variable para la carga de pruebas
    ' NBL: 25/6/2009
    ' ********************************************************************************
    Public Sub CerrarCargaPruebasgaPruebas()

        CerrarConexion()

    End Sub

    ' ********************************************************************************
    ' getDescripcionPruebas
    ' Desc: Función que devuelve la abrv y la descripción de una prueba y de las muestras
    ' NBL: 21/10/2009
    ' ********************************************************************************
    Public Sub getAbrvDescripcionPruebas(ByRef pstrCodigo As String, ByRef pstrAbrv As String, ByRef pstrDescripcion As String, _
                                                                ByRef pstrCodigoMuestra As String, ByRef pstrAbrvMuestra As String, ByRef pstrDescripcionMuestra As String, _
                                                                ByRef pstrTipoPrueba As String)

        mobjPruebas.getAbrvDescripcionPruebas(pstrCodigo, pstrAbrv, pstrDescripcion, _
                                                                    pstrCodigoMuestra, pstrAbrvMuestra, pstrDescripcionMuestra, pstrTipoPrueba, _
                                                                    Me.mobjCnnAccess, Me.mobjCnnSqlServer, Me.mobjCnnODBC)

    End Sub

    ' ********************************************************************************
    ' getDescripcionPruebas
    ' Desc: Función que devuelve la abrv y la descripción de una prueba y de las muestras
    ' NBL: 21/10/2009
    ' ********************************************************************************
    Public Sub getAbrvDescripcionPruebas(ByRef pstrCodigo As String, ByRef pstrAbrv As String, ByRef pstrDescripcion As String, _
                                                                ByRef pstrCodigoMuestra As String, ByRef pstrAbrvMuestra As String, ByRef pstrDescripcionMuestra As String, _
                                                                ByRef pstrTipoPrueba As String, _
                                                                ByVal pbolMicro As Boolean, ByVal pbolperfil As Boolean)

        mobjPruebas.getAbrvDescripcionPruebas(pstrCodigo, pstrAbrv, pstrDescripcion, _
                                                                    pstrCodigoMuestra, pstrAbrvMuestra, pstrDescripcionMuestra, pstrTipoPrueba, _
                                                                    Me.mobjCnnAccess, Me.mobjCnnSqlServer, Me.mobjCnnODBC, pbolMicro, pbolPerfil)

    End Sub

    ' ********************************************************************************
    ' getDescripcionPruebas
    ' Desc: Función que devuelve la abrv y la descripción de una prueba
    ' NBL: 25/6/2009
    ' ********************************************************************************
    Public Sub getAbrvDescripcionPruebas(ByRef pstrCodigo As String, ByRef pstrAbrv As String, ByRef pstrDescripcion As String, _
                                                                ByRef pstrTipoPrueba As String)

        mobjPruebas.getAbrvDescripcionPruebas(pstrCodigo, pstrAbrv, pstrDescripcion, pstrTipoPrueba, Me.mobjCnnAccess, Me.mobjCnnSqlServer, Me.mobjCnnODBC)

    End Sub

    Sub New()

        Dim mstrNombreArchivoConfig As String = clsUtil.DLLPath(True) + "Config.xml"
        clsUtil.CargarConfiguracion(mobjConfig, mstrNombreArchivoConfig)

        Dim mstrNombreArchivoConfigBBDD As String = ""
        If Me.mobjConfig.Conexion = BBDD.Local Then
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigLocalBBDD.xml"
        Else
            mstrNombreArchivoConfigBBDD = clsUtil.DLLPath(True) + "ConfigRemotaBBDD.xml"
        End If
        clsUtil.CargarConfiguracionBBDD(mobjConfigBBDD, mstrNombreArchivoConfigBBDD)

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
            MessageBox.Show("No se ha podido conectar a la BBDD" + vbCrLf + ex.Message, "Consulta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
                    If mobjCnnAccess.State <> ConnectionState.Closed Then
                        mobjCnnAccess.Close()
                    End If
                    'mobjCnnAccess.Dispose()
                    mobjCnnAccess = Nothing
                Else
                    If mobjCnnSqlServer.State <> ConnectionState.Closed Then mobjCnnSqlServer.Close()
                    'mobjCnnSqlServer.Dispose()
                    mobjCnnSqlServer = Nothing
                End If
            Else
                If mobjCnnODBC.State <> ConnectionState.Closed Then
                    mobjCnnODBC.Close()
                End If
                'mobjCnnODBC.Dispose()
                mobjCnnODBC = Nothing
            End If
            Return True
        Catch ex As Exception
            'MessageBox.Show("Error al cerrar la conexion a BBDD" + vbCrLf + ex.Message, "Consulta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ************************************************************************
    ' sqlGetMedicoLocal
    ' Desc: Función que crea el sql local para médicos
    ' NBL: 27/02/2007
    ' ************************************************************************
    Private Function sqlGetMedicoLocal() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Medicos.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Medicos.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMedicoLocal_1
    ' Desc: Función que crea el sql local para médicos 
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetMedicoLocal_1() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA", _
                                                                        Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Medicos.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA", _
                                                                        Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Medicos.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDestinoLocal
    ' Desc: Función que crea el sql local para destinos
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetDestinoLocal() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then

            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Destinos.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Destinos.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDestinoLocal_1
    ' Desc: Función que crea el sql local para destinos
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetDestinoLocal_1() As String

        Dim lstrSQL As String = ""

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Destinos.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA ORDER BY {1}", _
                                                                        Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                        Me.mobjConfigBBDD.Destinos.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetNHCbyNHUSA
    ' Desc: Función que crea el sql local para busqueda de NHC por NHUSA
    ' NBL: 17/11/2009
    ' ************************************************************************
    Private Function sqlGetNHCbyNHUSA(ByVal pstrRutaINI As String, ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String = ""
        Dim lobjINI As New clsINI

        Dim lstrTabla As String = lobjINI.IniGet(pstrRutaINI, "Parameters", "Tabla_NHC_NUHSA", "")
        Dim lstrNHC As String = lobjINI.IniGet(pstrRutaINI, "Parameters", "CampoNHC", "")
        Dim lstrNHUSA As String = lobjINI.IniGet(pstrRutaINI, "Parameters", "CampoNUHSA", "")

        If lstrTabla.Trim.Length <> 0 And lstrNHC.Trim.Length <> 0 And lstrNHUSA.Trim.Length <> 0 Then

            If clsUtil.NoCaseSensitive <> 1 Then
                lstrSQL = String.Format("SELECT {0} FROM {1} WHERE {2} = '{3}'", lstrNHC, lstrTabla, lstrNHUSA, pstrTextoBusqueda)
            Else
                lstrSQL = String.Format("SELECT {0} FROM {1} WHERE UCASE({2}) = '{3}'", lstrNHC, lstrTabla, lstrNHUSA, pstrTextoBusqueda)
            End If

            Return lstrSQL

        Else

            Return ""

        End If

    End Function

    ' ************************************************************************
    ' sqlGetNHUSAbyNHC
    ' Desc: Función que crea el sql local para busqueda de NHUSA por NHC
    ' NBL: 17/11/2009
    ' ************************************************************************
    Private Function sqlGetNHUSAbyNHC(ByVal pstrRutaINI As String, ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String = ""
        Dim lobjINI As New clsINI

        Dim lstrTabla As String = lobjINI.IniGet(pstrRutaINI, "General", "Tabla", "")
        Dim lstrNHC As String = lobjINI.IniGet(pstrRutaINI, "General", "NHC", "")
        Dim lstrNHUSA As String = lobjINI.IniGet(pstrRutaINI, "General", "NHUSA", "")

        If lstrTabla.Trim.Length <> 0 And lstrNHC.Trim.Length <> 0 And lstrNHUSA.Trim.Length <> 0 Then

            If clsUtil.NoCaseSensitive <> 1 Then
                lstrSQL = String.Format("SELECT {0} FROM {1} WHERE {2} = '{3}'", lstrNHUSA, lstrTabla, lstrNHC, pstrTextoBusqueda)
            Else
                lstrSQL = String.Format("SELECT {0} FROM {1} WHERE UCASE({2}) = '{3}'", lstrNHUSA, lstrTabla, lstrNHC, pstrTextoBusqueda)
            End If

            Return lstrSQL

        Else

            Return ""

        End If

    End Function

    ' ************************************************************************
    ' sqlGetCorrelacionLocal
    ' Desc: Función que crea el sql local para correlaciones
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function sqlGetCorrelacionLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDD.Correlaciones
            If clsUtil.NoCaseSensitive <> 1 Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} FROM {11} WHERE {0} = @BUSQUEDA", _
                                .ID, .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, .CODIGO_ORIGEN, _
                                .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, .CODIGO_GRUPO_FACTURACION, .CODIGO_CARGO, .TABLA)
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} FROM {11} WHERE UCASE({0}) = @BUSQUEDA", _
                                .ID, .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, .CODIGO_ORIGEN, _
                                .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, .CODIGO_GRUPO_FACTURACION, .CODIGO_CARGO, .TABLA)
            End If
        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetCorrelacionRemota
    ' Desc: Función que crea el sql remota para correlaciones
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function sqlGetCorrelacionRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDD.Correlaciones

            If clsUtil.NoCaseSensitive <> 1 Then
                lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} FROM {11} WHERE {0} = '{12}'", _
                                .ID, .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, .CODIGO_ORIGEN, _
                                .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, .CODIGO_GRUPO_FACTURACION, .CODIGO_CARGO, .TABLA, _
                                pstrTextoBusqueda)
            Else
                lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} FROM {11} WHERE UCASE({0}) = '{12}'", _
                                .ID, .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, .CODIGO_ORIGEN, _
                                .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, .CODIGO_GRUPO_FACTURACION, .CODIGO_CARGO, .TABLA, _
                                pstrTextoBusqueda.ToUpper)
            End If
        End With

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetHCLocal
    ' Desc: Función que crea el SQL local para HC
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function sqlGetHCLocal(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDD.HistoriasClinicas

            lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} FROM {14}", _
                            .ID, .NUM_HISTORIA, .APELLIDOS, .NOMBRE, .NUM_SS, .COD_SEXO, .FECHA_NACIMIENTO, _
                            .DNI, .DIRECCION, .POBLACION, .COD_PROVINCIA, .COD_POSTAL, .TELEFONO, .NOMBRE_COMPLETO, .TABLA)

            If Regex.IsMatch(pstrTextoBusqueda, Me.mobjConfig.reHC) Then
                lstrSQL += String.Format(" WHERE {0} = @BUSQUEDA", .NUM_HISTORIA)
            Else
                lstrSQL += String.Format(" WHERE {0} LIKE @BUSQUEDA", .NOMBRE_COMPLETO)
            End If

        End With



        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetHCRemota
    ' Desc: Función que crea el SQL remota para HC
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function sqlGetHCRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDD.HistoriasClinicas

            lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} FROM {14}", _
                            .ID, .NUM_HISTORIA, .APELLIDOS, .NOMBRE, .NUM_SS, .COD_SEXO, .FECHA_NACIMIENTO, _
                            .DNI, .DIRECCION, .POBLACION, .COD_PROVINCIA, .COD_POSTAL, .TELEFONO, .NOMBRE_COMPLETO, .TABLA)

            If Regex.IsMatch(pstrTextoBusqueda, Me.mobjConfig.reHC) Then
                lstrSQL += String.Format(" WHERE {0} = '{1}'", .NUM_HISTORIA, pstrTextoBusqueda)
            Else
                ' NBL: 30/5/2007
                ' Las consultas de historia clínica siempre serán Empieza por
                'If Me.mobjConfig.TipoConsulta = 0 Then
                lstrSQL += String.Format(" WHERE {0} LIKE '{1}'", .NOMBRE_COMPLETO, pstrTextoBusqueda + "%")
                'Else
                'lstrSQL += String.Format(" WHERE {0} LIKE '{1}'", .NOMBRE_COMPLETO, "%" + pstrTextoBusqueda + "%")
                'End If
            End If
        End With

        ' OJO
        'MessageBox.Show(lstrSQL)

        Return lstrSQL

    End Function


    ' ************************************************************************
    ' sqlGetTipoLocal
    ' Desc: Función que crea el sql local para Tipos
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetTipoLocal() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Tipo.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Tipo.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDiagnosticoLocal
    ' Desc: Función que crea el sql local para diagnósticos
    ' NBL: 12/06/2007
    ' ************************************************************************
    Private Function sqlGetDiagnosticoLocal() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDiagnosticoLocal_1
    ' Desc: Función que crea el sql local para un diagnóstico
    ' NBL: 13/06/2007
    ' ************************************************************************
    Private Function sqlGetDiagnosticoLocal_1() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetServicioLocal_1
    ' Desc: Función que crea el sql local para Servicios
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetServicioLocal_1() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetServicioLocal_1
    ' Desc: Función que crea el sql local para Servicios
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetTipoLocal_1() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Tipo.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetOrigenLocal
    ' Desc: Función que crea el sql local para Origenes
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetOrigenLocal() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Origenes.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Origenes.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetOrigenLocal_1
    ' Desc: Función que crea el sql local para un Origen
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetOrigenLocal_1() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = @BUSQUEDA ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMedicoRemota
    ' Desc: Función que crea el sql remoto para médicos
    ' NBL: 27/02/2007
    ' ************************************************************************
    Private Function sqlGetMedicoRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Medicos.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Medicos.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetMedicoRemota_1
    ' Desc: Función que crea el sql remoto para un médico
    ' NBL: 27/02/2007
    ' ************************************************************************
    Private Function sqlGetMedicoRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Medicos.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Medicos.CODIGO, Me.mobjConfigBBDD.Medicos.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Medicos.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function


    ' ************************************************************************
    ' sqlGetDestinoRemota
    ' Desc: Función que crea el sql remoto para destinos
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetDestinoRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Destinos.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Destinos.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL
    End Function

    ' ************************************************************************
    ' sqlGetDestinoRemota_1
    ' Desc: Función que crea el sql remoto para  un destino
    ' NBL: 21/0/2007
    ' ************************************************************************
    Private Function sqlGetDestinoRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Destinos.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Destinos.CODIGO, Me.mobjConfigBBDD.Destinos.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Destinos.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL
    End Function

    ' ************************************************************************
    ' sqlGetServicioRemota
    ' Desc: Función que crea el sql remoto para Servicios
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetServicioRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDiagnosticoRemota
    ' Desc: Función que crea el sql remoto para diagnósticos
    ' NBL: 12/06/2007
    ' ************************************************************************
    Private Function sqlGetDiagnosticoRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetDiagnosticoRemota_1
    ' Desc: Función que crea el sql remoto para un diagnóstico
    ' NBL: 13/06/2007
    ' ************************************************************************
    Private Function sqlGetDiagnosticoRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        'If Me.mobjConfig.TipoLocal = 0 Then
        '    lstrTextoBusqueda = pstrTextoBusqueda + "%"
        'Else
        '    lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        'End If

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Diagnostico.CODIGO, Me.mobjConfigBBDD.Diagnostico.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Diagnostico.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetServicioRemota_1
    ' Desc: Función que crea el sql remoto para Servicios
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetServicioRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Servicios.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Servicios.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetTipoRemota_1
    ' Desc: Función que crea el sql remoto para Tipos
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetTipoRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Tipo.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Tipo.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetTipoRemota
    ' Desc: Función que crea el sql remoto para Tipoes
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetTipoRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Tipo.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Tipo.CODIGO, Me.mobjConfigBBDD.Tipo.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Tipo.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetOrigenRemota
    ' Desc: Función que crea el sql remoto para Origenes
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetOrigenRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        If Me.mobjConfig.TipoLocal = 0 Then
            lstrTextoBusqueda = pstrTextoBusqueda + "%"
        Else
            lstrTextoBusqueda = "%" + pstrTextoBusqueda + "%"
        End If

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE '{3}' OR {1} LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE '{3}' OR UCASE({1}) LIKE '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' sqlGetOrigenRemota_1
    ' Desc: Función que crea el sql remoto para un Origen
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function sqlGetOrigenRemota_1(ByVal pstrTextoBusqueda As String) As String

        Dim lstrTextoBusqueda As String

        lstrTextoBusqueda = pstrTextoBusqueda

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA, lstrTextoBusqueda)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) = '{3}' ORDER BY {1}", _
                                                                    Me.mobjConfigBBDD.Origenes.CODIGO, Me.mobjConfigBBDD.Origenes.NOMBRE, _
                                                                    Me.mobjConfigBBDD.Origenes.TABLA, lstrTextoBusqueda.ToUpper)
        End If

        Return lstrSQL
    End Function

    ' ************************************************************************
    ' getMedicoLocalAccess
    ' Desc: Función de busqueda de médicos por BBDD local y access
    ' NBL: 27/02/2007
    ' ************************************************************************
    Private Function getMedicoLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandMedicos As New OleDbCommand(sqlGetMedicoLocal(), mobjCnnAccess)
        Dim lobjParamMedicos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMedicos.Parameters.Add(lobjParamMedicos)
        Dim lobjDataReaderMedicos As OleDbDataReader = lobjCommandMedicos.ExecuteReader()

        Dim larlstMedicos As New ArrayList()

        While lobjDataReaderMedicos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMedicos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMedicos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstMedicos.Add(lobjListViewItem)
        End While

        lobjDataReaderMedicos.Close()

        Return larlstMedicos

    End Function

    ' ************************************************************************
    ' getMedicoLocalAccess_1
    ' Desc: Función de busqueda un médico
    ' NBL: 21/5/2007
    ' ************************************************************************
    Private Function getMedicoLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandMedicos As New OleDbCommand(sqlGetMedicoLocal_1(), mobjCnnAccess)
        Dim lobjParamMedicos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommandMedicos.Parameters.Add(lobjParamMedicos)
        Dim lobjDataReaderMedicos As OleDbDataReader = lobjCommandMedicos.ExecuteReader()
        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderMedicos.Read()

            lstrCodigo = CType(lobjDataReaderMedicos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderMedicos.GetValue(1), String)
            End If

        End While

        lobjDataReaderMedicos.Close()

        Return lstrCodigo + "|" + lstrNombre

    End Function

    ' ************************************************************************
    ' getCorrelacionLocalAccess
    ' Desc: Función de busqueda de correlaciones por BBDD local y access
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function getCorrelacionLocalAccess(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommand As New OleDbCommand(sqlGetCorrelacionLocal(), mobjCnnAccess)
        Dim lobjParameter As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        lobjParameter.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommand.Parameters.Add(lobjParameter)

        Dim lobjDataReader As OleDbDataReader = lobjCommand.ExecuteReader()

        Dim lstrBResultado As New StringBuilder

        If lobjDataReader.Read() Then
            For lintContador As Integer = 1 To 10
                If lobjDataReader.IsDBNull(lintContador) Then
                    lstrBResultado.Append("")
                Else
                    lstrBResultado.Append(CType(lobjDataReader.GetValue(lintContador), String))
                End If
                If lintContador < 10 Then lstrBResultado.Append("|")
            Next
        End If

        lobjDataReader.Close()

        Return lstrBResultado.ToString()

    End Function

    ' ************************************************************************
    ' getCorrelacionLocalServer
    ' Desc: Función de busqueda de correlaciones por BBDD local y server
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function getCorrelacionLocalServer(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommand As New SqlCommand(sqlGetCorrelacionLocal(), mobjCnnSqlServer)
        Dim lobjParameter As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        lobjParameter.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommand.Parameters.Add(lobjParameter)

        Dim lobjDataReader As SqlDataReader = lobjCommand.ExecuteReader()

        Dim lstrBResultado As New StringBuilder

        If lobjDataReader.Read() Then
            For lintContador As Integer = 1 To 10
                If lobjDataReader.IsDBNull(lintContador) Then
                    lstrBResultado.Append("")
                Else
                    lstrBResultado.Append(CType(lobjDataReader.GetValue(lintContador), String))
                End If
                If lintContador < 10 Then lstrBResultado.Append("|")
            Next
        End If

        lobjDataReader.Close()

        Return lstrBResultado.ToString()

    End Function

    ' ************************************************************************
    ' getCorrelacionRemota
    ' Desc: Función de busqueda de correlaciones por BBDD remota
    ' NBL: 2/03/2007
    ' ************************************************************************
    Private Function getCorrelacionRemota(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommand As New OdbcCommand(sqlGetCorrelacionRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReader As OdbcDataReader = lobjCommand.ExecuteReader()

        Dim lstrBResultado As New StringBuilder

        If lobjDataReader.Read() Then
            For lintContador As Integer = 1 To 10
                If lobjDataReader.IsDBNull(lintContador) Then
                    lstrBResultado.Append("")
                Else
                    lstrBResultado.Append(CType(lobjDataReader.GetValue(lintContador), String))
                End If
                If lintContador < 10 Then lstrBResultado.Append("|")
            Next
        End If

        lobjDataReader.Close()

        Return lstrBResultado.ToString()

    End Function

    ' ************************************************************************
    ' getDestinoLocalAccess
    ' Desc: Función de busqueda de destinos por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getDestinoLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDestinos As New OleDbCommand(sqlGetDestinoLocal(), mobjCnnAccess)
        Dim lobjParamDestinos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandDestinos.Parameters.Add(lobjParamDestinos)
        Dim lobjDataReaderDestinos As OleDbDataReader = lobjCommandDestinos.ExecuteReader()

        Dim larlstDestinos As New ArrayList()

        While lobjDataReaderDestinos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDestinos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDestinos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDestinos.Add(lobjListViewItem)
        End While

        lobjDataReaderDestinos.Close()

        Return larlstDestinos

    End Function

    ' ************************************************************************
    ' getNHCbyNHUSAAccess
    ' Desc: Función de busqueda de NHUSA por NHC por BBDD local y access
    ' pintResultado 1 NHC, 2 NHUSA
    ' NBL: 17/11/2009
    ' ************************************************************************
    Private Function getNHCNHUSAAccess(ByVal pstrBusqueda As String, ByVal pstrRutaINI As String, ByVal pintResultado As Integer) As String

        Dim lobjCommand As OleDbCommand
        If pintResultado = 1 Then
            lobjCommand = New OleDbCommand(sqlGetNHCbyNHUSA(pstrRutaINI, pstrBusqueda), mobjCnnAccess)
        Else
            lobjCommand = New OleDbCommand(sqlGetNHUSAbyNHC(pstrRutaINI, pstrBusqueda), mobjCnnAccess)
        End If

        Dim lobjDataReader As OleDbDataReader = lobjCommand.ExecuteReader()

        Dim lstrResultado As String = ""

        If lobjDataReader.Read() Then
            If Not lobjDataReader.IsDBNull(0) Then
                lstrResultado = lobjDataReader.GetString(0)
            End If
        End If

        lobjDataReader.Close()

        Return lstrResultado

    End Function

    ' ************************************************************************
    ' getNHCNHUSASQLRemota
    ' Desc: Función de busqueda de NHUSA por NHC por BBDD local y access
    ' pintResultado 1 NHC, 2 NHUSA
    ' NBL: 17/11/2009
    ' ************************************************************************
    Private Function getNHCNHUSARemota(ByVal pstrBusqueda As String, ByVal pstrRutaINI As String, ByVal pintResultado As Integer) As String

        Dim lobjCommand As OdbcCommand
        If pintResultado = 1 Then
            lobjCommand = New OdbcCommand(sqlGetNHCbyNHUSA(pstrRutaINI, pstrBusqueda), mobjCnnODBC)
        Else
            lobjCommand = New OdbcCommand(sqlGetNHUSAbyNHC(pstrRutaINI, pstrBusqueda), mobjCnnODBC)
        End If

        Dim lobjDataReader As OdbcDataReader = lobjCommand.ExecuteReader()

        Dim lstrResultado As String = ""

        If lobjDataReader.Read() Then
            If Not lobjDataReader.IsDBNull(0) Then
                lstrResultado = lobjDataReader.GetString(0)
            End If
        End If

        lobjDataReader.Close()

        Return lstrResultado

    End Function

    ' ************************************************************************
    ' getNHCNHUSASQLSERVER
    ' Desc: Función de busqueda de NHUSA por NHC por BBDD local y access
    ' pintResultado 1 NHC, 2 NHUSA
    ' NBL: 17/11/2009
    ' ************************************************************************
    Private Function getNHCNHUSAServer(ByVal pstrBusqueda As String, ByVal pstrRutaINI As String, ByVal pintResultado As Integer) As String

        Dim lobjCommand As SqlCommand
        If pintResultado = 1 Then
            lobjCommand = New SqlCommand(sqlGetNHCbyNHUSA(pstrRutaINI, pstrBusqueda), mobjCnnSqlServer)
        Else
            lobjCommand = New SqlCommand(sqlGetNHUSAbyNHC(pstrRutaINI, pstrBusqueda), mobjCnnSqlServer)
        End If

        Dim lobjDataReader As SqlDataReader = lobjCommand.ExecuteReader()

        Dim lstrResultado As String = ""

        If lobjDataReader.Read() Then
            If Not lobjDataReader.IsDBNull(0) Then
                lstrResultado = lobjDataReader.GetString(0)
            End If
        End If

        lobjDataReader.Close()

        Return lstrResultado

    End Function

    ' ************************************************************************
    ' getDestinoLocalAccess_1
    ' Desc: Función de busqueda de un destino por BBDD local y access
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function getDestinoLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandDestinos As New OleDbCommand(sqlGetDestinoLocal_1(), mobjCnnAccess)
        Dim lobjParamDestinos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandDestinos.Parameters.Add(lobjParamDestinos)
        Dim lobjDataReaderDestinos As OleDbDataReader = lobjCommandDestinos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDestinos.Read()
            lstrCodigo = CType(lobjDataReaderDestinos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDestinos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDestinos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getDiagnosticoLocalServer_1
    ' Desc: Función de busqueda de un diagnóstico por BBDD local y server
    ' NBL: 13/06/2007
    ' ************************************************************************
    Private Function getDiagnosticoLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandDiagnostico As New SqlCommand(sqlGetDiagnosticoLocal_1, mobjCnnSqlServer)
        Dim lobjParamDiagnostico As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommandDiagnostico.Parameters.Add(lobjParamDiagnostico)

        Dim lobjDataReaderDiagnostico As SqlDataReader = lobjCommandDiagnostico.ExecuteReader()
        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDiagnostico.Read()
            lstrCodigo = CType(lobjDataReaderDiagnostico.GetValue(0), String)
            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDiagnostico.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDiagnostico.Close()

        Return lstrCodigo + "|" + lstrNombre

    End Function

    ' ************************************************************************
    ' getHCLocalAccess
    ' Desc: Función de busqueda de HC por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getHCLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandHC As New OleDbCommand(sqlGetHCLocal(pstrTextoBusqueda), mobjCnnAccess)
        Dim lobjParamHC As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        If lobjCommandHC.CommandText.IndexOf("= @BUSQUEDA") >= 0 Then
            lobjParamHC.Value = pstrTextoBusqueda
        Else
            ' NBL: 30/5/2007
            ' Comento estas lineas porque la búsqueda será empieza por siempre para el diálogo de historias clínicas
            'If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamHC.Value = pstrTextoBusqueda + "%"
            'Else
            'lobjParamHC.Value = "%" + pstrTextoBusqueda + "%"
            'End If
        End If

        lobjCommandHC.Parameters.Add(lobjParamHC)

        Dim lobjDataReaderHC As OleDbDataReader = lobjCommandHC.ExecuteReader()
        Dim larlstHC As New ArrayList

        While lobjDataReaderHC.Read

            ' Creamos el item y le damos el valor de HC
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderHC.GetValue(1), String))

            ' Apellidos
            If lobjDataReaderHC.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(2), String))
            End If

            ' Nombre
            If lobjDataReaderHC.IsDBNull(3) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(3), String))
            End If

            ' Fecha Nacimiento
            ' NBL: 8/6/2007. Formateamos la fecha en formato dd/mm/yyyy
            Dim lstrFecha As String = ""
            If lobjDataReaderHC.IsDBNull(6) Then
                lstrFecha = ""
            Else
                lstrFecha = CType(lobjDataReaderHC.GetValue(6), String)
                If lstrFecha.Length = 8 Then
                    lstrFecha = lstrFecha.Substring(6, 2) + "/" + lstrFecha.Substring(4, 2) + "/" + lstrFecha.Substring(0, 4)
                End If
            End If
            lobjListViewItem.SubItems.Add(lstrFecha)

            Dim lstrBTodo As New StringBuilder
            ' Hacemos un bucle por todos las columnas de la consulta
            For lintcontador As Integer = 0 To 13
                If lobjDataReaderHC.IsDBNull(lintcontador) Then
                    lstrBTodo.Append("")
                Else
                    lstrBTodo.Append(CType(lobjDataReaderHC.GetValue(lintcontador), String))
                End If
                If lintcontador < 13 Then lstrBTodo.Append("|")
            Next

            lobjListViewItem.Tag = lstrBTodo.ToString()

            larlstHC.Add(lobjListViewItem)

        End While

        lobjDataReaderHC.Close()

        Return larlstHC

    End Function

    ' ************************************************************************
    ' getHCLocalServer
    ' Desc: Función de busqueda de HC por BBDD local y server
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getHCLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandHC As New SqlCommand(sqlGetHCLocal(pstrTextoBusqueda), mobjCnnSqlServer)
        Dim lobjParamHC As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If lobjCommandHC.CommandText.IndexOf("= @BUSQUEDA") >= 0 Then
            lobjParamHC.Value = pstrTextoBusqueda
        Else
            ' NBL: 30/5/2007
            ' Las consultas de historias clínicas siempre serán Empieza por
            'If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamHC.Value = pstrTextoBusqueda + "%"
            'Else
            'lobjParamHC.Value = "%" + pstrTextoBusqueda + "%"
            'End If
        End If
        lobjCommandHC.Parameters.Add(lobjParamHC)

        Dim lobjDataReaderHC As SqlDataReader = lobjCommandHC.ExecuteReader()
        Dim larlstHC As New ArrayList

        While lobjDataReaderHC.Read

            ' Creamos el item y le damos el valor de HC
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderHC.GetValue(1), String))

            ' Apellidos
            If lobjDataReaderHC.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(2), String))
            End If

            ' Nombre
            If lobjDataReaderHC.IsDBNull(3) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(3), String))
            End If

            ' Fecha Nacimiento
            ' NBL: 8/6/2007. Formateamos la fecha en formato dd/mm/yyyy
            Dim lstrFecha As String = ""
            If lobjDataReaderHC.IsDBNull(6) Then
                lstrFecha = ""
            Else
                lstrFecha = CType(lobjDataReaderHC.GetValue(6), String)
                If lstrFecha.Length = 8 Then
                    lstrFecha = lstrFecha.Substring(6, 2) + "/" + lstrFecha.Substring(4, 2) + "/" + lstrFecha.Substring(0, 4)
                End If
            End If
            lobjListViewItem.SubItems.Add(lstrFecha)

            Dim lstrBTodo As New StringBuilder
            ' Hacemos un bucle por todos las columnas de la consulta
            For lintcontador As Integer = 0 To 13
                If lobjDataReaderHC.IsDBNull(lintcontador) Then
                    lstrBTodo.Append("")
                Else
                    lstrBTodo.Append(CType(lobjDataReaderHC.GetValue(lintcontador), String))
                End If
                If lintcontador < 13 Then lstrBTodo.Append("|")
            Next

            lobjListViewItem.Tag = lstrBTodo.ToString()

            larlstHC.Add(lobjListViewItem)

        End While

        lobjDataReaderHC.Close()

        Return larlstHC

    End Function

    ' ************************************************************************
    ' getHCRemota
    ' Desc: Función de busqueda de HC por BBDD remota
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getHCRemota(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandHC As New OdbcCommand(sqlGetHCRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderHC As OdbcDataReader = lobjCommandHC.ExecuteReader()
        Dim larlstHC As New ArrayList
        ' MessageBox.Show("Consulta finalizada")

        While lobjDataReaderHC.Read

            ' Creamos el item y le damos el valor de HC
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderHC.GetValue(1), String))

            ' Apellidos
            If lobjDataReaderHC.IsDBNull(2) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(2), String))
            End If

            ' Nombre
            If lobjDataReaderHC.IsDBNull(3) Then
                lobjListViewItem.SubItems.Add("")
            Else
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderHC.GetValue(3), String))
            End If

            ' Fecha Nacimiento
            ' NBL: 8/6/2007. Formateamos la fecha en formato dd/mm/yyyy
            Dim lstrFecha As String = ""
            If lobjDataReaderHC.IsDBNull(6) Then
                lstrFecha = ""
            Else
                lstrFecha = CType(lobjDataReaderHC.GetValue(6), String)
                If lstrFecha.Length = 8 Then
                    lstrFecha = lstrFecha.Substring(6, 2) + "/" + lstrFecha.Substring(4, 2) + "/" + lstrFecha.Substring(0, 4)
                End If
            End If
            lobjListViewItem.SubItems.Add(lstrFecha)

            Dim lstrBTodo As New StringBuilder
            ' Hacemos un bucle por todos las columnas de la consulta
            For lintcontador As Integer = 0 To 13
                If lobjDataReaderHC.IsDBNull(lintcontador) Then
                    lstrBTodo.Append("")
                Else
                    lstrBTodo.Append(CType(lobjDataReaderHC.GetValue(lintcontador), String))
                End If
                If lintcontador < 13 Then lstrBTodo.Append("|")
            Next

            lobjListViewItem.Tag = lstrBTodo.ToString()

            larlstHC.Add(lobjListViewItem)

        End While

        lobjDataReaderHC.Close()
        Return larlstHC

    End Function

    ' ************************************************************************
    ' getServicioLocalAccess_1
    ' Desc: Función de busqueda de un servicio por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getServicioLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandServicios As New OleDbCommand(sqlGetServicioLocal_1(), mobjCnnAccess)
        Dim lobjParamServicios As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandServicios.Parameters.Add(lobjParamServicios)
        Dim lobjDataReaderServicios As OleDbDataReader = lobjCommandServicios.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderServicios.Read()
            lstrCodigo = CType(lobjDataReaderServicios.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderServicios.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderServicios.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getTipoLocalAccess_1
    ' Desc: Función de busqueda de un servicio por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getTipoLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandTipos As New OleDbCommand(sqlGetTipoLocal(), mobjCnnAccess)
        Dim lobjParamTipos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandTipos.Parameters.Add(lobjParamTipos)
        Dim lobjDataReaderTipos As OleDbDataReader = lobjCommandTipos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderTipos.Read()
            lstrCodigo = CType(lobjDataReaderTipos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderTipos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderTipos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getDiagnosticoLocalAccess
    ' Desc: Función de busqueda de diagnósticos por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getDiagnosticoLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDiagnostico As New OleDbCommand(sqlGetDiagnosticoLocal, mobjCnnAccess)
        Dim lobjParamDiagnostico As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If

        lobjCommandDiagnostico.Parameters.Add(lobjParamDiagnostico)
        Dim lobjDataReaderDiagnostico As OleDbDataReader = lobjCommandDiagnostico.ExecuteReader

        Dim larlstDiagnostico As New ArrayList()

        While lobjDataReaderDiagnostico.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDiagnostico.GetValue(0), String))

            ' Ponemos la descripción del diagnóstico
            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDiagnostico.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDiagnostico.Add(lobjListViewItem)
        End While

        lobjDataReaderDiagnostico.Close()

        Return larlstDiagnostico

    End Function

    ' ************************************************************************
    ' getDiagnosticoLocalAccess_1
    ' Desc: Función de busqueda de un diagnóstico por BBDD local y access
    ' NBL: 13/06/2007
    ' ************************************************************************
    Private Function getDiagnosticoLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandDiagnostico As New OleDbCommand(sqlGetDiagnosticoLocal_1, mobjCnnAccess)
        Dim lobjParamDiagnostico As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommandDiagnostico.Parameters.Add(lobjParamDiagnostico)

        Dim lobjDataReaderDiagnostico As OleDbDataReader = lobjCommandDiagnostico.ExecuteReader()
        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDiagnostico.Read()
            lstrCodigo = CType(lobjDataReaderDiagnostico.GetValue(0), String)
            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDiagnostico.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDiagnostico.Close()

        Return lstrCodigo + "|" + lstrNombre

    End Function

    ' ************************************************************************
    ' getDiagnosticoLocalServer
    ' Desc: Función de busqueda de diagnósticos por BBDD local y server
    ' NBL: 12/06/2007
    ' ************************************************************************
    Private Function getDiagnosticoLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDiagnostico As New SqlCommand(sqlGetDiagnosticoLocal, mobjCnnSqlServer)
        Dim lobjParamDiagnostico As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamDiagnostico.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If

        lobjCommandDiagnostico.Parameters.Add(lobjParamDiagnostico)
        Dim lobjDataReaderDiagnostico As SqlDataReader = lobjCommandDiagnostico.ExecuteReader

        Dim larlstDiagnostico As New ArrayList()

        While lobjDataReaderDiagnostico.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDiagnostico.GetValue(0), String))

            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDiagnostico.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDiagnostico.Add(lobjListViewItem)
        End While

        lobjDataReaderDiagnostico.Close()

        Return larlstDiagnostico

    End Function

    ' ************************************************************************
    ' getServicioLocalAccess
    ' Desc: Función de busqueda de Servicios por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getServicioLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandServicios As New OleDbCommand(sqlGetServicioLocal(), mobjCnnAccess)
        Dim lobjParamServicios As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandServicios.Parameters.Add(lobjParamServicios)
        Dim lobjDataReaderServicios As OleDbDataReader = lobjCommandServicios.ExecuteReader()

        Dim larlstServicios As New ArrayList()

        While lobjDataReaderServicios.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderServicios.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderServicios.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstServicios.Add(lobjListViewItem)
        End While

        lobjDataReaderServicios.Close()

        Return larlstServicios

    End Function

    ' ************************************************************************
    ' getTipoLocalAccess
    ' Desc: Función de busqueda de Tipos por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getTipoLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandTipos As New OleDbCommand(sqlGetTipoLocal(), mobjCnnAccess)
        Dim lobjParamTipos As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandTipos.Parameters.Add(lobjParamTipos)
        Dim lobjDataReaderTipos As OleDbDataReader = lobjCommandTipos.ExecuteReader()

        Dim larlstTipos As New ArrayList()

        While lobjDataReaderTipos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderTipos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderTipos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstTipos.Add(lobjListViewItem)
        End While

        lobjDataReaderTipos.Close()

        Return larlstTipos

    End Function

    ' ************************************************************************
    ' getOrigenLocalAccess
    ' Desc: Función de busqueda de Origenes por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getOrigenLocalAccess(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandOrigenes As New OleDbCommand(sqlGetOrigenLocal(), mobjCnnAccess)
        Dim lobjParamOrigenes As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandOrigenes.Parameters.Add(lobjParamOrigenes)
        Dim lobjDataReaderOrigenes As OleDbDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim larlstOrigenes As New ArrayList()

        While lobjDataReaderOrigenes.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderOrigenes.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderOrigenes.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstOrigenes.Add(lobjListViewItem)
        End While

        lobjDataReaderOrigenes.Close()

        Return larlstOrigenes

    End Function

    ' ************************************************************************
    ' getOrigenLocalAccess_1
    ' Desc: Función de busqueda de un origen por BBDD local y access
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function getOrigenLocalAccess_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandOrigenes As New OleDbCommand(sqlGetOrigenLocal_1(), mobjCnnAccess)
        Dim lobjParamOrigenes As New OleDbParameter("@BUSQUEDA", OleDbType.WChar)

        lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandOrigenes.Parameters.Add(lobjParamOrigenes)
        Dim lobjDataReaderOrigenes As OleDbDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderOrigenes.Read()
            lstrCodigo = CType(lobjDataReaderOrigenes.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderOrigenes.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderOrigenes.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getMedicoLocalServer_1
    ' Desc: Función de busqueda de un médico por BBDD local y server
    ' NBL: 26/02/2007
    ' ************************************************************************
    Private Function getMedicoLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandMedicos As New SqlCommand(sqlGetMedicoLocal_1(), mobjCnnSqlServer)
        Dim lobjParamMedicos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandMedicos.Parameters.Add(lobjParamMedicos)
        Dim lobjDataReaderMedicos As SqlDataReader = lobjCommandMedicos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderMedicos.Read()
            lstrCodigo = CType(lobjDataReaderMedicos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderMedicos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderMedicos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getMedicoLocalServer
    ' Desc: Función de busqueda de médicos por BBDD local y server
    ' NBL: 26/02/2007
    ' ************************************************************************
    Private Function getMedicoLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandMedicos As New SqlCommand(sqlGetMedicoLocal(), mobjCnnSqlServer)
        Dim lobjParamMedicos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamMedicos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandMedicos.Parameters.Add(lobjParamMedicos)
        Dim lobjDataReaderMedicos As SqlDataReader = lobjCommandMedicos.ExecuteReader()

        Dim larlstMedicos As New ArrayList()

        While lobjDataReaderMedicos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMedicos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMedicos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstMedicos.Add(lobjListViewItem)
        End While

        lobjDataReaderMedicos.Close()

        Return larlstMedicos

    End Function

    ' ************************************************************************
    ' getDestinoLocalServer
    ' Desc: Función de busqueda de destinos por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getDestinoLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDestinos As New SqlCommand(sqlGetDestinoLocal(), mobjCnnSqlServer)
        Dim lobjParamDestinos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandDestinos.Parameters.Add(lobjParamDestinos)
        Dim lobjDataReaderDestinos As SqlDataReader = lobjCommandDestinos.ExecuteReader()

        Dim larlstDestinos As New ArrayList()

        While lobjDataReaderDestinos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDestinos.GetValue(0), String))

            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDestinos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDestinos.Add(lobjListViewItem)
        End While

        lobjDataReaderDestinos.Close()

        Return larlstDestinos

    End Function

    ' ************************************************************************
    ' getDestinoLocalServer_1
    ' Desc: Función de busqueda de un destino por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getDestinoLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandDestinos As New SqlCommand(sqlGetDestinoLocal_1(), mobjCnnSqlServer)
        Dim lobjParamDestinos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamDestinos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)
        lobjCommandDestinos.Parameters.Add(lobjParamDestinos)
        Dim lobjDataReaderDestinos As SqlDataReader = lobjCommandDestinos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDestinos.Read()
            lstrCodigo = CType(lobjDataReaderDestinos.GetValue(0), String)

            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDestinos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDestinos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' sqlGetServicioLocal
    ' Desc: Función que crea el sql local para Servicios
    ' NBL: 28/02/2007
    ' ************************************************************************
    Private Function sqlGetServicioLocal() As String

        Dim lstrSQL As String

        If clsUtil.NoCaseSensitive <> 1 Then
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE {0} LIKE @BUSQUEDA OR {1} LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA)
        Else
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2} WHERE UCASE({0}) LIKE @BUSQUEDA OR UCASE({1}) LIKE @BUSQUEDA ORDER BY {1}", _
                                                                                Me.mobjConfigBBDD.Servicios.CODIGO, Me.mobjConfigBBDD.Servicios.NOMBRE, _
                                                                                Me.mobjConfigBBDD.Servicios.TABLA)
        End If

        Return lstrSQL

    End Function

    ' ************************************************************************
    ' getServicioLocalServer
    ' Desc: Función de busqueda de Servicios por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getServicioLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandServicios As New SqlCommand(sqlGetServicioLocal(), mobjCnnSqlServer)
        Dim lobjParamServicios As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandServicios.Parameters.Add(lobjParamServicios)
        Dim lobjDataReaderServicios As SqlDataReader = lobjCommandServicios.ExecuteReader()

        Dim larlstServicios As New ArrayList()

        While lobjDataReaderServicios.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderServicios.GetValue(0), String))

            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderServicios.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstServicios.Add(lobjListViewItem)
        End While

        lobjDataReaderServicios.Close()

        Return larlstServicios

    End Function


    ' ************************************************************************
    ' getTipoLocalServer
    ' Desc: Función de busqueda de Tipos por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getTipoLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandTipos As New SqlCommand(sqlGetTipoLocal(), mobjCnnSqlServer)
        Dim lobjParamTipos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandTipos.Parameters.Add(lobjParamTipos)
        Dim lobjDataReaderTipos As SqlDataReader = lobjCommandTipos.ExecuteReader()

        Dim larlstTipos As New ArrayList()

        While lobjDataReaderTipos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderTipos.GetValue(0), String))

            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderTipos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstTipos.Add(lobjListViewItem)
        End While

        lobjDataReaderTipos.Close()

        Return larlstTipos

    End Function

    ' ************************************************************************
    ' getServicioLocalServer_1
    ' Desc: Función de busqueda de un Servicio por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getServicioLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandServicios As New SqlCommand(sqlGetServicioLocal_1(), mobjCnnSqlServer)
        Dim lobjParamServicios As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamServicios.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandServicios.Parameters.Add(lobjParamServicios)
        Dim lobjDataReaderServicios As SqlDataReader = lobjCommandServicios.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderServicios.Read()
            lstrCodigo = CType(lobjDataReaderServicios.GetValue(0), String)

            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderServicios.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderServicios.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getTipoLocalServer_1
    ' Desc: Función de busqueda de un Tipo por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getTipoLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandTipos As New SqlCommand(sqlGetTipoLocal_1(), mobjCnnSqlServer)
        Dim lobjParamTipos As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamTipos.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandTipos.Parameters.Add(lobjParamTipos)
        Dim lobjDataReaderTipos As SqlDataReader = lobjCommandTipos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderTipos.Read()
            lstrCodigo = CType(lobjDataReaderTipos.GetValue(0), String)

            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderTipos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderTipos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getOrigenLocalServer_1
    ' Desc: Función de busqueda de un Origen por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getOrigenLocalServer_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandOrigenes As New SqlCommand(sqlGetOrigenLocal_1(), mobjCnnSqlServer)
        Dim lobjParamOrigenes As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)

        lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda, pstrTextoBusqueda.ToUpper)

        lobjCommandOrigenes.Parameters.Add(lobjParamOrigenes)
        Dim lobjDataReaderOrigenes As SqlDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderOrigenes.Read()
            lstrCodigo = CType(lobjDataReaderOrigenes.GetValue(0), String)

            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderOrigenes.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderOrigenes.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getOrigenLocalServer
    ' Desc: Función de busqueda de Origenes por BBDD local y server
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getOrigenLocalServer(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandOrigenes As New SqlCommand(sqlGetOrigenLocal(), mobjCnnSqlServer)
        Dim lobjParamOrigenes As New SqlParameter("@BUSQUEDA", SqlDbType.VarChar)
        If Me.mobjConfig.TipoConsulta = 0 Then
            lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, pstrTextoBusqueda + "%", pstrTextoBusqueda.ToUpper + "%")
        Else
            lobjParamOrigenes.Value = IIf(clsUtil.NoCaseSensitive <> 1, "%" + pstrTextoBusqueda + "%", "%" + pstrTextoBusqueda.ToUpper + "%")
        End If
        lobjCommandOrigenes.Parameters.Add(lobjParamOrigenes)
        Dim lobjDataReaderOrigenes As SqlDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim larlstOrigenes As New ArrayList()

        While lobjDataReaderOrigenes.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderOrigenes.GetValue(0), String))

            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderOrigenes.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstOrigenes.Add(lobjListViewItem)
        End While

        lobjDataReaderOrigenes.Close()

        Return larlstOrigenes

    End Function

    ' ************************************************************************
    ' getMedicoRemoto_1
    ' Desc: Función de busqueda de medico remoto
    ' NBL: 22/05/2007
    ' ************************************************************************
    Private Function getMedicoRemoto_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandMedicos As New OdbcCommand(sqlGetMedicoRemota_1(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderMedicos As OdbcDataReader = lobjCommandMedicos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderMedicos.Read()
            lstrCodigo = CType(lobjDataReaderMedicos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderMedicos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If

        End While

        lobjDataReaderMedicos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getMedicoRemoto
    ' Desc: Función de busqueda de medico remoto
    ' NBL: 26/02/2007
    ' ************************************************************************
    Private Function getMedicoRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandMedicos As New OdbcCommand(sqlGetMedicoRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderMedicos As OdbcDataReader = lobjCommandMedicos.ExecuteReader()

        Dim larlstMedicos As New ArrayList()

        While lobjDataReaderMedicos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderMedicos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderMedicos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderMedicos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstMedicos.Add(lobjListViewItem)
        End While

        lobjDataReaderMedicos.Close()

        Return larlstMedicos

    End Function

    ' ************************************************************************
    ' getDestinoRemoto
    ' Desc: Función de busqueda de destino remoto
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getDestinoRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDestinos As New OdbcCommand(sqlGetDestinoRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderDestinos As OdbcDataReader = lobjCommandDestinos.ExecuteReader()

        Dim larlstDestinos As New ArrayList()

        While lobjDataReaderDestinos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDestinos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDestinos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDestinos.Add(lobjListViewItem)
        End While

        lobjDataReaderDestinos.Close()

        Return larlstDestinos

    End Function

    ' ************************************************************************
    ' getDestinoRemoto_1
    ' Desc: Función de busqueda de un destino remoto
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function getDestinoRemoto_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandDestinos As New OdbcCommand(sqlGetDestinoRemota_1(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderDestinos As OdbcDataReader = lobjCommandDestinos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDestinos.Read()
            lstrCodigo = CType(lobjDataReaderDestinos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderDestinos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDestinos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDestinos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getDiagnosticoRemoto
    ' Desc: Función de busqueda de diagnóstico remoto
    ' NBL: 12/06/2007
    ' ************************************************************************
    Private Function getDiagnosticoRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandDiagnostico As New OdbcCommand(sqlGetDiagnosticoRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderDiagnostico As OdbcDataReader = lobjCommandDiagnostico.ExecuteReader()

        Dim larlstDiagnostico As New ArrayList()

        While lobjDataReaderDiagnostico.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderDiagnostico.GetValue(0), String))

            'Ponemos la descripción del perfil
            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderDiagnostico.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstDiagnostico.Add(lobjListViewItem)
        End While

        lobjDataReaderDiagnostico.Close()

        Return larlstDiagnostico

    End Function

    ' ************************************************************************
    ' getDiagnosticoRemoto_1
    ' Desc: Función de busqueda de diagnóstico remoto
    ' NBL: 13/06/2007
    ' ************************************************************************
    Private Function getDiagnosticoRemoto_1(ByVal pstrtextobusqueda As String) As String

        Dim lobjCommandDiagnostico As New OdbcCommand(sqlGetDiagnosticoRemota_1(pstrtextobusqueda), mobjCnnODBC)
        Dim lobjDataReaderDiagnostico As OdbcDataReader = lobjCommandDiagnostico.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderDiagnostico.Read()
            lstrCodigo = CType(lobjDataReaderDiagnostico.GetValue(0), String)
            If Not lobjDataReaderDiagnostico.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderDiagnostico.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderDiagnostico.Close()

        Return lstrCodigo + "|" + lstrNombre

    End Function

    ' ************************************************************************
    ' getTipoRemoto
    ' Desc: Función de busqueda de Tipo remoto
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getTipoRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandTipos As New OdbcCommand(sqlGetTipoRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderTipos As OdbcDataReader = lobjCommandTipos.ExecuteReader()

        Dim larlstTipos As New ArrayList()

        While lobjDataReaderTipos.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderTipos.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderTipos.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstTipos.Add(lobjListViewItem)
        End While

        lobjDataReaderTipos.Close()

        Return larlstTipos

    End Function

    ' ************************************************************************
    ' getServicioRemoto
    ' Desc: Función de busqueda de Servicio remoto
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getServicioRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandServicios As New OdbcCommand(sqlGetServicioRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderServicios As OdbcDataReader = lobjCommandServicios.ExecuteReader()

        Dim larlstServicios As New ArrayList()

        While lobjDataReaderServicios.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderServicios.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderServicios.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstServicios.Add(lobjListViewItem)
        End While

        lobjDataReaderServicios.Close()

        Return larlstServicios

    End Function

    ' ************************************************************************
    ' getServicioRemoto_1
    ' Desc: Función de busqueda de un Servicio remoto
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function getServicioRemoto_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandServicios As New OdbcCommand(sqlGetServicioRemota_1(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderServicios As OdbcDataReader = lobjCommandServicios.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderServicios.Read()
            lstrCodigo = CType(lobjDataReaderServicios.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderServicios.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderServicios.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderServicios.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getTipoRemoto_1
    ' Desc: Función de busqueda de un Servicio remoto
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function getTipoRemoto_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandTipos As New OdbcCommand(sqlGetTipoRemota_1(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderTipos As OdbcDataReader = lobjCommandTipos.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderTipos.Read()
            lstrCodigo = CType(lobjDataReaderTipos.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderTipos.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderTipos.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderTipos.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' getOrigenRemoto
    ' Desc: Función de busqueda de Origen remoto
    ' NBL: 1/03/2007
    ' ************************************************************************
    Private Function getOrigenRemoto(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim lobjCommandOrigenes As New OdbcCommand(sqlGetOrigenRemota(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderOrigenes As OdbcDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim larlstOrigenes As New ArrayList()

        While lobjDataReaderOrigenes.Read()
            Dim lobjListViewItem As ListViewItem = New ListViewItem(CType(lobjDataReaderOrigenes.GetValue(0), String))

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lobjListViewItem.SubItems.Add(CType(lobjDataReaderOrigenes.GetValue(1), String))
            Else
                lobjListViewItem.SubItems.Add("")
            End If
            lobjListViewItem.Name = lobjListViewItem.Text

            ' Lo metemos dentro del array
            larlstOrigenes.Add(lobjListViewItem)
        End While

        lobjDataReaderOrigenes.Close()

        Return larlstOrigenes

    End Function

    ' ************************************************************************
    ' getOrigenRemoto_1
    ' Desc: Función de busqueda de Origen remoto
    ' NBL: 21/05/2007
    ' ************************************************************************
    Private Function getOrigenRemoto_1(ByVal pstrTextoBusqueda As String) As String

        Dim lobjCommandOrigenes As New OdbcCommand(sqlGetOrigenRemota_1(pstrTextoBusqueda), mobjCnnODBC)
        Dim lobjDataReaderOrigenes As OdbcDataReader = lobjCommandOrigenes.ExecuteReader()

        Dim lstrCodigo As String = ""
        Dim lstrNombre As String = ""

        While lobjDataReaderOrigenes.Read()
            lstrCodigo = CType(lobjDataReaderOrigenes.GetValue(0), String)

            ' Ponemos la descripción del perfil
            If Not lobjDataReaderOrigenes.IsDBNull(1) Then
                lstrNombre = CType(lobjDataReaderOrigenes.GetValue(1), String)
            Else
                lstrNombre = ""
            End If
        End While

        lobjDataReaderOrigenes.Close()

        Return lstrCodigo & "|" & lstrNombre

    End Function

    ' ************************************************************************
    ' BuscaMedico
    ' Desc: Función de busqueda de médicos
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaMedico(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getMedicoLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getMedicoLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getMedicoRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de médicos" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaMedico_1
    ' Desc: Función de busqueda de un médico
    ' NBL: 22/05/2007
    ' ************************************************************************
    Public Function BuscaMedico_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getMedicoLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getMedicoLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota

                Resultado = getMedicoRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de médicos 1" + vbCrLf + ex.Message, "Consulta médicos 1", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' getNHCbyNHUSA
    ' Desc: Función de busqueda de NHC a partir de NHUSA,
    ' MODIFICACIÓN SOLO PARA EL HOSPITAL VIRGEN DEL ROCÍO
    ' NBL: 16/11/2009
    ' ************************************************************************
    Public Function getNHCbyNHUSA(ByVal pstrNHUSA As String, ByVal pstrRutaINI As String) As String

        Dim Resultado As String = ""

        Try
            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getNHCNHUSAAccess(pstrNHUSA, pstrRutaINI, 1)
                Else
                    ' SQL Server
                    Resultado = getNHCNHUSAServer(pstrNHUSA, pstrRutaINI, 1)
                End If
            Else
                ' BBDD remota
                Resultado = getNHCNHUSARemota(pstrNHUSA, pstrRutaINI, 1)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' getNHUSAbyNHC
    ' Desc: Función de busqueda de NHUSA a partir de NHC
    ' MODIFICACIÓN SOLO PARA EL HOSPITAL VIRGEN DEL ROCÍO
    ' NBL: 16/11/2009
    ' ************************************************************************
    Public Function getNHUSAbyNHC(ByVal pstrNHC As String, ByVal pstrRutaINI As String) As String

        Dim Resultado As String = ""

        Try
            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing
            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getNHCNHUSAAccess(pstrNHC, pstrRutaINI, 2)
                Else
                    ' SQL Server
                    Resultado = getNHCNHUSAServer(pstrNHC, pstrRutaINI, 2)
                End If
            Else
                ' BBDD remota
                Resultado = getNHCNHUSARemota(pstrNHC, pstrRutaINI, 2)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaDestino_1
    ' Desc: Función de busqueda de un destino
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function BuscaDestino_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getDestinoLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getDestinoLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getDestinoRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de destinos" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaDestino
    ' Desc: Función de busqueda de destinos
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaDestino(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getDestinoLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getDestinoLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getDestinoRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de destinos" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaServicio_1
    ' Desc: Función de busqueda de un servicio
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function BuscaServicio_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getServicioLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getServicioLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getServicioRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de Servicios" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaTipo_1
    ' Desc: Función de busqueda de un Tipo
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function BuscaTipo_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getTipoLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getTipoLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getTipoRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de Tipos" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaServicio_1
    ' Desc: Función de busqueda de un servicio
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function BuscaTipo1_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getTipoLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getTipoLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getTipoRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de Tipos" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaDiagnostico
    ' Desc: Función de busqueda de diagnóstico
    ' NBL: 12/06/2007
    ' ************************************************************************
    Public Function BuscaDiagnostico(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getDiagnosticoLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL Server
                    Resultado = getDiagnosticoLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getDiagnosticoRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de diagnósticos" + vbCrLf + ex.Message, "Consulta diagnósticos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaDiagnostico_1
    ' Desc: Función de busqueda de diagnóstico_1
    ' NBL: 13/06/2007
    ' ************************************************************************
    Public Function BuscaDiagnostico_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getDiagnosticoLocalAccess_1(pstrTextoBusqueda)
                Else
                    'SQL Server
                    Resultado = getDiagnosticoLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getDiagnosticoRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de diagnósticos" + vbCrLf + ex.Message, "Consulta diagnósticos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaServicio
    ' Desc: Función de busqueda de Servicios
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaServicio(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getServicioLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getServicioLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getServicioRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de Servicios" + vbCrLf + ex.Message, "Consulta servicios", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function


    ' ************************************************************************
    ' BuscaTipo
    ' Desc: Función de busqueda de Tipos
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaTipo(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getTipoLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getTipoLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getTipoRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de Tipos" + vbCrLf + ex.Message, "Consulta Tipos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaHC
    ' Desc: Función de busqueda de historia clínica
    ' NBL: 1/03/2007
    ' ************************************************************************
    Public Function BuscaHC(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getHCLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getHCLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getHCRemota(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de historia clínica" + vbCrLf + ex.Message, "Consulta HC", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            'MessageBox.Show(ex.Source)
            'MessageBox.Show(ex.Message)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaOrigen_1
    ' Desc: Función de busqueda de un Origen
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function BuscaOrigen_1(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return ""

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getOrigenLocalAccess_1(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getOrigenLocalServer_1(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getOrigenRemoto_1(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            'MessageBox.Show("Error en la consulta de Origenes" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

    ' ************************************************************************
    ' BuscaOrigen
    ' Desc: Función de busqueda de Origenes
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaOrigen(ByVal pstrTextoBusqueda As String) As ArrayList

        Dim Resultado As ArrayList = Nothing

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return Nothing

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getOrigenLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getOrigenLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getOrigenRemoto(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()
            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de Origenes" + vbCrLf + ex.Message, "Consulta médicos", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return Nothing

        End Try

    End Function

    ' ************************************************************************
    ' BuscaCorrelacion
    ' Desc: Función de busqueda de correlación
    ' NBL: 26/02/2007
    ' ************************************************************************
    Public Function BuscaCorrelacion(ByVal pstrTextoBusqueda As String) As String

        Dim Resultado As String = ""

        Try

            ' Abrimos la conexión para hacer la consulta
            If Not AbrirConexion() Then Return ""

            If Me.mobjConfig.Conexion = BBDD.Local Then
                ' BBDD local
                If Me.mobjConfig.TipoLocal = 0 Then
                    ' Access
                    Resultado = getCorrelacionLocalAccess(pstrTextoBusqueda)
                Else
                    ' SQL server
                    Resultado = getCorrelacionLocalServer(pstrTextoBusqueda)
                End If
            Else
                ' BBDD remota
                Resultado = getCorrelacionRemota(pstrTextoBusqueda)
            End If

            ' Cerramos la conexión
            CerrarConexion()

            Return Resultado

        Catch ex As Exception

            CerrarConexion()
            MessageBox.Show("Error en la consulta de correlaciones" + vbCrLf + ex.Message, "Consulta correlación", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return ""

        End Try

    End Function

End Class


