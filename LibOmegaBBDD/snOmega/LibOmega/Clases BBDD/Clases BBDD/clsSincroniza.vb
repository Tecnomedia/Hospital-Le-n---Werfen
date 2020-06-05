Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.Odbc
Imports System.Windows.Forms

' Creo el enum para saber de donde viene el elemento insertado
Public Enum TablaSincro
    Bioquimica
    Muestra
    Microbiologia
    MicroMuestra
    PerfilBioquimica
    PerfilMicro
    Medicos
    HistoriasClinicas
    Correlaciones
    Origenes
    Servicios
    Destinos
    Diagnosticos
End Enum

Public Class clsSincroniza

    ' Instancias de configuración
    Private mobjConfig As clsConfig
    Private mobjConfigBBDDLocal As clsConfigBBDD
    Private mobjConfigBBDDRemota As clsConfigBBDD

    ' Ponemos las tres tipos de conexiones que puede haber en este formulario
    Private mobjCnnAccess As OleDbConnection
    Private mobjCnnSqlServer As SqlConnection
    Private mobjCnnODBC As OdbcConnection

    ' Creamos el Delegado del evento
    Public Delegate Sub InsertarRegistroEventHandler(ByVal sender As Object, ByVal e As clsSincronizaEventsArgs)
    Public Event InsertarRegistro As InsertarRegistroEventHandler

    ' Metodo que lanza el evento
    Protected Overridable Sub OnInsertarRegistro(ByVal e As clsSincronizaEventsArgs)
        RaiseEvent InsertarRegistro(Me, e)
    End Sub

    Public Sub New()

        Dim mstrNombreArchivoConfig As String = clsUtil.DLLPath(True) + "Config.xml"
        clsUtil.CargarConfiguracion(mobjConfig, mstrNombreArchivoConfig)

        Dim mstrNombreArchivoConfigBBDDLocal As String = clsUtil.DLLPath(True) + "ConfigLocalBBDD.xml"
        Dim mstrNombreArchivoConfigBBDDRemota As String = clsUtil.DLLPath(True) + "ConfigRemotaBBDD.xml"

        'MsgBox(mstrNombreArchivoConfigBBDDLocal)
        'MsgBox(mstrNombreArchivoConfigBBDDRemota)

        clsUtil.CargarConfiguracionBBDD(mobjConfigBBDDLocal, mstrNombreArchivoConfigBBDDLocal)
        clsUtil.CargarConfiguracionBBDD(mobjConfigBBDDRemota, mstrNombreArchivoConfigBBDDRemota)

    End Sub

    ' ************************************************************************
    ' AbrirConexion
    ' Desc: En la sincronización ha de abrir tanto la BBDD remota como la local
    ' NBL: 19/01/2007
    ' ************************************************************************
    Private Function AbrirConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota

            If Me.mobjConfig.TipoLocal = 0 Then
                ' La BBDD es Access
                mobjCnnAccess = New OleDbConnection(Me.mobjConfig.cnLocal)
                mobjCnnAccess.Open()
            Else
                ' SQL Server
                mobjCnnSqlServer = New SqlConnection(Me.mobjConfig.cnLocal)
                mobjCnnSqlServer.Open()
            End If

            ' ODBC hacia OMEGA
            mobjCnnODBC = New OdbcConnection("DSN=" + Me.mobjConfig.DSNRemota)
            mobjCnnODBC.Open()

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
    Private Function CerrarConexion() As Boolean

        Try
            ' Primero tenemos que saber si la conexión es local o remota

            If Me.mobjConfig.TipoLocal = 0 Then
                If mobjCnnAccess.State <> ConnectionState.Closed Then mobjCnnAccess.Close()
                mobjCnnAccess.Dispose()
                mobjCnnAccess = Nothing
            Else
                If mobjCnnSqlServer.State <> ConnectionState.Closed Then mobjCnnSqlServer.Close()
                mobjCnnSqlServer = Nothing
                mobjCnnSqlServer.Dispose()
            End If

            If mobjCnnODBC.State <> ConnectionState.Closed Then mobjCnnODBC.Close()
            mobjCnnODBC.Dispose()
            mobjCnnODBC = Nothing

            Call System.GC.Collect()

            Return True
        Catch ex As Exception
            MessageBox.Show("Error al cerrar la conexion a BBDD" + vbCrLf + ex.Message, "Consulta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' NumeroRegistroTotal
    ' Desc: Función que devuelve el número de registros que tiene una tabla que pasemos por parámetro
    ' NBL: 9/3/2007
    ' ********************************************************************************
    Private Function NumeroRegistroTotal(ByVal pstrNombreTabla As String) As Integer

        Try
            Dim lstrSQL As String = String.Format("SELECT COUNT(*) FROM {0}", pstrNombreTabla)
            Dim lobjCommand As New OdbcCommand(lstrSQL, mobjCnnODBC)
            Dim lintNumeroTotal As Integer = lobjCommand.ExecuteScalar()
            Return lintNumeroTotal
        Catch ex As Exception
            Return 0
        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaPruebasBioquimica
    ' Desc: Función que sincroniza las tablas locales de pruebas de bioquimica con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaPruebasBioquimica() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaBioquimicaLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Bioquimica.TABLA)

        Try
            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectPruebasBioquimicaRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim ldblCodigo As Double
                Dim lstrAbrv As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    ldblCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrAbrv = ""
                Else
                    lstrAbrv = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                If lobjDataReaderRemoto.IsDBNull(2) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(2), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    InsertarPruebaBioquimicaLocal(ldblCodigo, lstrAbrv, lstrNom)
                    lintContador += 1
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Bioquimica)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()
            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaPerfilBioquimica
    ' Desc: Función que sincroniza las tablas locales de perfiles de bioquímica con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaPerfilBioquimica() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaPerfilBioquimicaLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.PerfilBioquimica.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectPerfilBioquimicaRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim ldblCodigo As Double
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    ldblCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarPerfilBioquimicaLocal(ldblCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.PerfilBioquimica)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()

            Return False


        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaPerfilMicro
    ' Desc: Función que sincroniza las tablas locales de perfiles de micro con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaPerfilMicro() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaPerfilMicroLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.PerfilMicro.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectPerfilMicroRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String = ""
                Dim lstrNom As String = ""
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), String)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarPerfilMicroLocal(lstrCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.PerfilMicro)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()

            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaCorrelaciones
    ' Desc: Función que sincroniza las tablas locales de correlaciones con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaCorrelaciones() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaCorrelacionesLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Correlaciones.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las correlaciones
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectCorrelacionesRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrID As String, lstrCAMPO As String, lstrDESENCADENANTE As String, lstrDOCTOR As String
                Dim lstrSERVICIO As String, lstrORIGEN As String, lstrDESTINO As String, lstrMOTIVO As String
                Dim lstrTIPO As String, lstrGRUPOFAC As String, lstrCARGO As String

                Dim lbolInsertOK As Boolean = False
                ' ID
                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrID = CType(lobjDataReaderRemoto.GetValue(0), String)
                    lbolInsertOK = True
                End If
                ' CAMPO
                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrCAMPO = ""
                Else
                    lstrCAMPO = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If
                ' DESENCADENANTE
                If lobjDataReaderRemoto.IsDBNull(2) Then
                    lstrDESENCADENANTE = ""
                Else
                    lstrDESENCADENANTE = CType(lobjDataReaderRemoto.GetValue(2), String)
                End If
                ' DOCTOR
                If lobjDataReaderRemoto.IsDBNull(3) Then
                    lstrDOCTOR = ""
                Else
                    lstrDOCTOR = CType(lobjDataReaderRemoto.GetValue(3), String)
                End If
                ' SERVICIO
                If lobjDataReaderRemoto.IsDBNull(4) Then
                    lstrSERVICIO = ""
                Else
                    lstrSERVICIO = CType(lobjDataReaderRemoto.GetValue(4), String)
                End If
                ' ORIGEN
                If lobjDataReaderRemoto.IsDBNull(5) Then
                    lstrORIGEN = ""
                Else
                    lstrORIGEN = CType(lobjDataReaderRemoto.GetValue(5), String)
                End If
                ' DESTINO
                If lobjDataReaderRemoto.IsDBNull(6) Then
                    lstrDESTINO = ""
                Else
                    lstrDESTINO = CType(lobjDataReaderRemoto.GetValue(6), String)
                End If
                ' MOTIVO
                If lobjDataReaderRemoto.IsDBNull(7) Then
                    lstrMOTIVO = ""
                Else
                    lstrMOTIVO = CType(lobjDataReaderRemoto.GetValue(7), String)
                End If
                ' TIPO
                If lobjDataReaderRemoto.IsDBNull(8) Then
                    lstrTIPO = ""
                Else
                    lstrTIPO = CType(lobjDataReaderRemoto.GetValue(8), String)
                End If
                ' GRUPOFAC
                If lobjDataReaderRemoto.IsDBNull(9) Then
                    lstrGRUPOFAC = ""
                Else
                    lstrGRUPOFAC = CType(lobjDataReaderRemoto.GetValue(9), String)
                End If
                ' CARGO
                If lobjDataReaderRemoto.IsDBNull(10) Then
                    lstrCARGO = ""
                Else
                    lstrCARGO = CType(lobjDataReaderRemoto.GetValue(10), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarCorrelacionLocal(lstrID, lstrCAMPO, lstrDESENCADENANTE, lstrDOCTOR, lstrSERVICIO, _
                                                                                            lstrORIGEN, lstrDESTINO, lstrMOTIVO, lstrTIPO, lstrGRUPOFAC, lstrCARGO)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Correlaciones)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaOrigenes
    ' Desc: Función que sincroniza las tablas locales de orígenes con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaOrigenes() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaOrigenesLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Origenes.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectOrigenesRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarOrigenLocal(lstrCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Origenes)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()

            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaDestinos
    ' Desc: Función que sincroniza las tablas locales de destinos con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaDestinos() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaDestinosLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Destinos.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectDestinosRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarDestinoLocal(lstrCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Destinos)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' sqlDeleteDiagnosticosLocal
    ' Desc: Función que devuelve el DELETE local de diagnósticos
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Private Function sqlDeleteDiagnosticosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Diagnostico
            lstrSQL = String.Format("DELETE FROM {0} ", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' SincronizaDiagnosticos
    ' Desc: Función que sincroniza las tablas locales de diagnósticos con las remotas
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Public Function SincronizaDiagnosticos() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaDiagnosticosLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Diagnostico.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectDiagnosticosRemotas, mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarDiagnosticoLocal(lstrCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Diagnosticos)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloqueado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaDiagnosticosLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 11/6/2007
    ' ********************************************************************************
    Private Function BorrarTablaDiagnosticosLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access
                Dim lobjCommand As New OleDbCommand(sqlDeleteDiagnosticosLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                'Server
                Dim lobjcommand As New SqlCommand(sqlDeleteDiagnosticosLocal(), mobjCnnSqlServer)
                lobjcommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de diagnósticos." + ex.Message, "BorrarTablaDiagnosticosLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaMedicos
    ' Desc: Función que sincroniza las tablas locales de médicos con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaMedicos() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaMedicosLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Medicos.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectMedicosRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), String)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción

                If lbolInsertOK Then
                    lintContador += 1
                    InsertarMedicoLocal(lstrCodigo, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Medicos)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()

            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaServicios
    ' Desc: Función que sincroniza las tablas locales de servicios con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaServicios() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaServiciosLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Servicios.TABLA)

        Try

            ' Hacemos la consulta que recoja todas los servicios
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectServiciosRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim lstrCodigo As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrCodigo = CType(lobjDataReaderRemoto.GetValue(0), String)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarServicioLocal(lstrCodigo, lstrNom)
                    Dim e As New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Servicios)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function


    ' ********************************************************************************
    ' SincronizaHC
    ' Desc: Función que sincroniza las tablas locales de historias clínicas con las remotas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Public Function SincronizaHC() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaHCLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.HistoriasClinicas.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectHCRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()
                Dim lstrID As String, lstrHC As String, lstrAPELLIDOS As String, lstrNOMBRE As String, lstrNUMEROSS As String
                Dim ldblCODIGOSEXO As Double, lstrFECHANAC As String, lstrDNI As String, lstrDIRECCION As String, lstrPOBLACION As String
                Dim lstrPROVINCIA As String, lstrCP As String, lstrTELEFONO As String, lstrNOMBRECOMPLETO As String

                Dim lbolInsertOK As Boolean = False
                ' ID
                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    lstrID = CType(lobjDataReaderRemoto.GetValue(0), String)
                    lbolInsertOK = True
                End If
                ' HC
                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrHC = ""
                Else
                    lstrHC = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If
                ' APELLIDOS
                If lobjDataReaderRemoto.IsDBNull(2) Then
                    lstrAPELLIDOS = ""
                Else
                    lstrAPELLIDOS = CType(lobjDataReaderRemoto.GetValue(2), String)
                End If
                ' NOMBRE
                If lobjDataReaderRemoto.IsDBNull(3) Then
                    lstrNOMBRE = ""
                Else
                    lstrNOMBRE = CType(lobjDataReaderRemoto.GetValue(3), String)
                End If
                ' NUMEROSS
                If lobjDataReaderRemoto.IsDBNull(4) Then
                    lstrNUMEROSS = ""
                Else
                    lstrNUMEROSS = CType(lobjDataReaderRemoto.GetValue(4), String)
                End If
                ' CODIGOSEXO
                If lobjDataReaderRemoto.IsDBNull(5) Then
                    ldblCODIGOSEXO = 0
                Else
                    ldblCODIGOSEXO = CType(lobjDataReaderRemoto.GetValue(5), Double)
                End If
                ' FECHANAC
                If lobjDataReaderRemoto.IsDBNull(6) Then
                    lstrFECHANAC = ""
                Else
                    lstrFECHANAC = CType(lobjDataReaderRemoto.GetValue(6), String)
                End If
                ' DNI
                If lobjDataReaderRemoto.IsDBNull(7) Then
                    lstrDNI = ""
                Else
                    lstrDNI = CType(lobjDataReaderRemoto.GetValue(7), String)
                End If
                ' DIRECCION
                If lobjDataReaderRemoto.IsDBNull(8) Then
                    lstrDIRECCION = ""
                Else
                    lstrDIRECCION = CType(lobjDataReaderRemoto.GetValue(8), String)
                End If
                ' POBLACIÓN
                If lobjDataReaderRemoto.IsDBNull(9) Then
                    lstrPOBLACION = ""
                Else
                    lstrPOBLACION = CType(lobjDataReaderRemoto.GetValue(9), String)
                End If
                ' PROVINCIA
                If lobjDataReaderRemoto.IsDBNull(10) Then
                    lstrPROVINCIA = ""
                Else
                    lstrPROVINCIA = CType(lobjDataReaderRemoto.GetValue(10), String)
                End If
                ' CP
                If lobjDataReaderRemoto.IsDBNull(11) Then
                    lstrCP = ""
                Else
                    lstrCP = CType(lobjDataReaderRemoto.GetValue(11), String)
                End If
                ' TELEFONO
                If lobjDataReaderRemoto.IsDBNull(12) Then
                    lstrTELEFONO = ""
                Else
                    lstrTELEFONO = CType(lobjDataReaderRemoto.GetValue(12), String)
                End If
                ' NOMBRECOMPLETO
                If lobjDataReaderRemoto.IsDBNull(13) Then
                    lstrNOMBRECOMPLETO = ""
                Else
                    lstrNOMBRECOMPLETO = CType(lobjDataReaderRemoto.GetValue(13), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarHCLocal(lstrID, lstrHC, lstrAPELLIDOS, lstrNOMBRE, lstrNUMEROSS, ldblCODIGOSEXO, lstrFECHANAC, lstrDNI, lstrDIRECCION, _
                                                                                lstrPOBLACION, lstrPROVINCIA, lstrCP, lstrTELEFONO, lstrNOMBRECOMPLETO)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.HistoriasClinicas)
                    OnInsertarRegistro(e)
                End If
                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaMuestras
    ' Desc: Función que sincroniza las tablas locales de muestras con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaMuestras() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaMuestraLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Muestra.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectMuestrasRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim ldblCodigo As Double
                Dim lstrAbrv As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    ldblCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrAbrv = ""
                Else
                    lstrAbrv = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                If lobjDataReaderRemoto.IsDBNull(2) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(2), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarMuestraLocal(ldblCodigo, lstrAbrv, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Muestra)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaMicroMuestras
    ' Desc: Función que sincroniza las tablas locales de micro-muestras con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaMicroMuestras() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaMicroMuestraLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.MicroMuestra.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectMicroMuestrasRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim ldblCodigoPrueba As Double
                Dim ldblCodigoMuestra As Double
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    ldblCodigoPrueba = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    ldblCodigoMuestra = 0
                Else
                    ldblCodigoMuestra = CType(lobjDataReaderRemoto.GetValue(1), Double)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarMicroMuestraLocal(ldblCodigoPrueba, ldblCodigoMuestra)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.MicroMuestra)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()
            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' SincronizaPruebasMicro
    ' Desc: Función que sincroniza las tablas locales de pruebas de micro con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaPruebasMicro() As Boolean

        ' Si no podemos abrir la conexión a la base de datos salimos
        If Not AbrirConexion() Then Return False

        ' Primero borramos la tabla de la BBDD local
        If Not BorrarTablaPruebasMicroLocal() Then Return False

        Dim lintContador As Integer = 0
        Dim lintNumeroTotal As Integer = NumeroRegistroTotal(Me.mobjConfigBBDDRemota.Microbiologia.TABLA)

        Try

            ' Hacemos la consulta que recoja todas las pruebas 
            Dim lobjCommandRemoto As New OdbcCommand(sqlSelectPruebasMicroRemotas(), mobjCnnODBC)
            Dim lobjDataReaderRemoto As OdbcDataReader = lobjCommandRemoto.ExecuteReader()

            ' Hacemos un bucle a través de todos los registros obtenidos para meterlos 
            While lobjDataReaderRemoto.Read()

                Dim ldblCodigo As Double
                Dim lstrAbrv As String
                Dim lstrNom As String
                Dim lbolInsertOK As Boolean = False

                If Not lobjDataReaderRemoto.IsDBNull(0) Then
                    ldblCodigo = CType(lobjDataReaderRemoto.GetValue(0), Double)
                    lbolInsertOK = True
                End If

                If lobjDataReaderRemoto.IsDBNull(1) Then
                    lstrAbrv = ""
                Else
                    lstrAbrv = CType(lobjDataReaderRemoto.GetValue(1), String)
                End If

                If lobjDataReaderRemoto.IsDBNull(2) Then
                    lstrNom = ""
                Else
                    lstrNom = CType(lobjDataReaderRemoto.GetValue(2), String)
                End If

                ' Llamamos a la rutina de inserción
                If lbolInsertOK Then
                    lintContador += 1
                    InsertarPruebasMicroLocal(ldblCodigo, lstrAbrv, lstrNom)
                    Dim e As clsSincronizaEventsArgs = New clsSincronizaEventsArgs(lintContador, lintNumeroTotal, TablaSincro.Microbiologia)
                    OnInsertarRegistro(e)
                End If

                ' Pongo aquí el DoEvents para que no se quede medio bloquedado
                Application.DoEvents()

            End While

            lobjDataReaderRemoto.Close()
            CerrarConexion()

            Return True

        Catch ex As Exception

            CerrarConexion()

            Return False

        End Try

    End Function

    ' ********************************************************************************
    ' InsertarPruebaBioquimicaLocal
    ' Desc: Rutina que inserta prueba de bioquímica a la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Sub InsertarPruebaBioquimicaLocal(ByVal pdblCodigo As Double, ByVal pstrAbrv As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertPruebasBioquimicaLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.Double)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New OleDbParameter("@ABRV", OleDbType.WChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertPruebasBioquimicaLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.Float)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New SqlParameter("@ABRV", SqlDbType.VarChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarMuestraLocal
    ' Desc: Rutina que inserta muestra a la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Sub InsertarMuestraLocal(ByVal pdblCodigo As Double, ByVal pstrAbrv As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertMuestrasLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.Double)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New OleDbParameter("@ABRV", OleDbType.WChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertMuestrasLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.Float)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New SqlParameter("@ABRV", SqlDbType.VarChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarPerfilBioquimicaLocal
    ' Desc: Rutina que inserta perfil bioquímica a la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Sub InsertarPerfilBioquimicaLocal(ByVal pdblCodigo As Double, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertPerfilBioquimicaLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.Double)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertPerfilBioquimicaLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.Float)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarOrigenLocal
    ' Desc: Rutina que inserta origen a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarOrigenLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertOrigenesLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertOrigenesLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarMedicoLocal
    ' Desc: Rutina que inserta médico a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarMedicoLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertMedicosLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertMedicosLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarCorrelacionLocal
    ' Desc: Rutina que inserta correlación a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarCorrelacionLocal(ByVal pstrID As String, ByVal pstrCAMPO As String, ByVal pstrDESENCADENANTE As String, _
                                                            ByVal pstrDOCTOR As String, ByVal pstrSERVICIO As String, ByVal pstrORIGEN As String, _
                                                            ByVal pstrDESTINO As String, ByVal pstrMOTIVO As String, ByVal pstrTIPO As String, _
                                                            ByVal pstrGRUPOFAC As String, ByVal pstrCARGO As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertCorrelacionesLocal(), mobjCnnAccess)
            ' ID
            Dim lobjParameterID As New OleDbParameter("@ID", OleDbType.WChar)
            lobjParameterID.Value = pstrID
            lobjCommand.Parameters.Add(lobjParameterID)
            'CAMPO
            Dim lobjParameterCAMPO As New OleDbParameter("@CAMPO", OleDbType.WChar)
            lobjParameterCAMPO.Value = pstrCAMPO
            lobjCommand.Parameters.Add(lobjParameterCAMPO)
            ' DESENCADENANTE
            Dim lobjParameterDESENCADENANTE As New OleDbParameter("@DESENCADENANTE", OleDbType.WChar)
            lobjParameterDESENCADENANTE.Value = pstrDESENCADENANTE
            lobjCommand.Parameters.Add(lobjParameterDESENCADENANTE)
            ' DOCTOR
            Dim lobjParameterDOCTOR As New OleDbParameter("@DOCTOR", OleDbType.WChar)
            lobjParameterDOCTOR.Value = pstrDOCTOR
            lobjCommand.Parameters.Add(lobjParameterDOCTOR)
            ' SERVICIO
            Dim lobjParameterSERVICIO As New OleDbParameter("@SERVICIO", OleDbType.WChar)
            lobjParameterSERVICIO.Value = pstrSERVICIO
            lobjCommand.Parameters.Add(lobjParameterSERVICIO)
            ' ORIGEN
            Dim lobjParameterORIGEN As New OleDbParameter("@ORIGEN", OleDbType.WChar)
            lobjParameterORIGEN.Value = pstrORIGEN
            lobjCommand.Parameters.Add(lobjParameterORIGEN)
            ' DESTINO
            Dim lobjParameterDESTINO As New OleDbParameter("@DESTINO", OleDbType.WChar)
            lobjParameterDESTINO.Value = pstrDESTINO
            lobjCommand.Parameters.Add(lobjParameterDESTINO)
            ' MOTIVO
            Dim lobjParameterMOTIVO As New OleDbParameter("@MOTIVO", OleDbType.WChar)
            lobjParameterMOTIVO.Value = pstrMOTIVO
            lobjCommand.Parameters.Add(lobjParameterMOTIVO)
            ' TIPO
            Dim lobjParameterTIPO As New OleDbParameter("@TIPO", OleDbType.WChar)
            lobjParameterTIPO.Value = pstrTIPO
            lobjCommand.Parameters.Add(lobjParameterTIPO)
            ' GRUPOFAC
            Dim lobjParameterGRUPOFAC As New OleDbParameter("@GRUPOFAC", OleDbType.WChar)
            lobjParameterGRUPOFAC.Value = pstrGRUPOFAC
            lobjCommand.Parameters.Add(lobjParameterGRUPOFAC)
            ' CARGO
            Dim lobjParameterCARGO As New OleDbParameter("@CARGO", OleDbType.WChar)
            lobjParameterCARGO.Value = pstrCARGO
            lobjCommand.Parameters.Add(lobjParameterCARGO)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertCorrelacionesLocal(), mobjCnnSqlServer)

            ' ID
            Dim lobjParameterID As New SqlParameter("@ID", SqlDbType.VarChar)
            lobjParameterID.Value = pstrID
            lobjCommand.Parameters.Add(lobjParameterID)
            'CAMPO
            Dim lobjParameterCAMPO As New SqlParameter("@CAMPO", SqlDbType.VarChar)
            lobjParameterCAMPO.Value = pstrCAMPO
            lobjCommand.Parameters.Add(lobjParameterCAMPO)
            ' DESENCADENANTE
            Dim lobjParameterDESENCADENANTE As New SqlParameter("@DESENCADENANTE", SqlDbType.VarChar)
            lobjParameterDESENCADENANTE.Value = pstrDESENCADENANTE
            lobjCommand.Parameters.Add(lobjParameterDESENCADENANTE)
            ' DOCTOR
            Dim lobjParameterDOCTOR As New SqlParameter("@DOCTOR", SqlDbType.VarChar)
            lobjParameterDOCTOR.Value = pstrDOCTOR
            lobjCommand.Parameters.Add(lobjParameterDOCTOR)
            ' SERVICIO
            Dim lobjParameterSERVICIO As New SqlParameter("@SERVICIO", SqlDbType.VarChar)
            lobjParameterSERVICIO.Value = pstrSERVICIO
            lobjCommand.Parameters.Add(lobjParameterSERVICIO)
            ' ORIGEN
            Dim lobjParameterORIGEN As New SqlParameter("@ORIGEN", SqlDbType.VarChar)
            lobjParameterORIGEN.Value = pstrORIGEN
            lobjCommand.Parameters.Add(lobjParameterORIGEN)
            ' DESTINO
            Dim lobjParameterDESTINO As New SqlParameter("@DESTINO", SqlDbType.VarChar)
            lobjParameterDESTINO.Value = pstrDESTINO
            lobjCommand.Parameters.Add(lobjParameterDESTINO)
            ' MOTIVO
            Dim lobjParameterMOTIVO As New SqlParameter("@MOTIVO", SqlDbType.VarChar)
            lobjParameterMOTIVO.Value = pstrMOTIVO
            lobjCommand.Parameters.Add(lobjParameterMOTIVO)
            ' TIPO
            Dim lobjParameterTIPO As New SqlParameter("@TIPO", SqlDbType.VarChar)
            lobjParameterTIPO.Value = pstrTIPO
            lobjCommand.Parameters.Add(lobjParameterTIPO)
            ' GRUPOFAC
            Dim lobjParameterGRUPOFAC As New SqlParameter("@GRUPOFAC", SqlDbType.VarChar)
            lobjParameterGRUPOFAC.Value = pstrGRUPOFAC
            lobjCommand.Parameters.Add(lobjParameterGRUPOFAC)
            ' CARGO
            Dim lobjParameterCARGO As New SqlParameter("@CARGO", SqlDbType.VarChar)
            lobjParameterCARGO.Value = pstrCARGO
            lobjCommand.Parameters.Add(lobjParameterCARGO)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarDiagnosticoLocal
    ' Desc: Rutina que inserta diagnósticos a la BBDD local
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Private Sub InsertarDiagnosticoLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertDiagnosticoLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertDiagnosticoLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarDestinoLocal
    ' Desc: Rutina que inserta destinos a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarDestinoLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertDestinosLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertDestinosLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarServicioLocal
    ' Desc: Rutina que inserta servicio a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarServicioLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertServiciosLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertServiciosLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarHCLocal
    ' Desc: Rutina que inserta HC a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarHCLocal(ByVal pstrID As String, ByVal pstrHC As String, ByVal pstrAPELLIDOS As String, ByVal pstrNOMBRE As String, _
                                                ByVal pstrNUMEROSS As String, ByVal pdblCODIGOSEXO As Double, ByVal pstrFECHANAC As String, _
                                                ByVal pstrDNI As String, ByVal pstrDIRECCION As String, ByVal pstrPOBLACION As String, _
                                                ByVal pstrPROVINCIA As String, ByVal pstrCP As String, ByVal pstrTELEFONO As String, _
                                                ByVal pstrNOMBRECOMPLETO As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertHCLocal(), mobjCnnAccess)
            ' ID
            Dim lobjParameterID As New OleDbParameter("@ID", OleDbType.WChar)
            lobjParameterID.Value = pstrID
            lobjCommand.Parameters.Add(lobjParameterID)
            ' Historias Clínicas
            Dim lobjParameterHC As New OleDbParameter("@HC", OleDbType.WChar)
            lobjParameterHC.Value = pstrHC
            lobjCommand.Parameters.Add(lobjParameterHC)
            ' Apellidos
            Dim lobjParameterAPELLIDOS As New OleDbParameter("@APELLIDOS", OleDbType.WChar)
            lobjParameterAPELLIDOS.Value = pstrAPELLIDOS
            lobjCommand.Parameters.Add(lobjParameterAPELLIDOS)
            ' Nombre
            Dim lobjParameterNOMBRE As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNOMBRE.Value = pstrNOMBRE
            lobjCommand.Parameters.Add(lobjParameterNOMBRE)
            ' Numero SS
            Dim lobjParameterNUMEROSS As New OleDbParameter("@NUMEROSS", OleDbType.WChar)
            lobjParameterNUMEROSS.Value = pstrNUMEROSS
            lobjCommand.Parameters.Add(lobjParameterNUMEROSS)
            ' Código sexo
            Dim lobjParameterCODIGOSEXO As New OleDbParameter("@CODIGOSEXO", OleDbType.Double)
            lobjParameterCODIGOSEXO.Value = pdblCODIGOSEXO
            lobjCommand.Parameters.Add(lobjParameterCODIGOSEXO)
            ' Fecha Nacimiento
            Dim lobjParameterFECHANAC As New OleDbParameter("@FECHANAC", OleDbType.WChar)
            lobjParameterFECHANAC.Value = pstrFECHANAC
            lobjCommand.Parameters.Add(lobjParameterFECHANAC)
            ' DNI
            Dim lobjParameterDNI As New OleDbParameter("@DNI", OleDbType.WChar)
            lobjParameterDNI.Value = pstrDNI
            lobjCommand.Parameters.Add(lobjParameterDNI)
            ' Dirección
            Dim lobjParameterDIRECCION As New OleDbParameter("@DIRECCION", OleDbType.WChar)
            lobjParameterDIRECCION.Value = pstrDIRECCION
            lobjCommand.Parameters.Add(lobjParameterDIRECCION)
            ' Población
            Dim lobjParameterPOBLACION As New OleDbParameter("@POBLACION", OleDbType.WChar)
            lobjParameterPOBLACION.Value = pstrPOBLACION
            lobjCommand.Parameters.Add(lobjParameterPOBLACION)
            ' Provincia
            Dim lobjParameterPROVINCIA As New OleDbParameter("@PROVINCIA", OleDbType.WChar)
            lobjParameterPROVINCIA.Value = pstrPROVINCIA
            lobjCommand.Parameters.Add(lobjParameterPROVINCIA)
            ' Código postal
            Dim lobjParameterCP As New OleDbParameter("@CP", OleDbType.WChar)
            lobjParameterCP.Value = pstrCP
            lobjCommand.Parameters.Add(lobjParameterCP)
            ' Teléfono
            Dim lobjParameterTELEFONO As New OleDbParameter("@TELEFONO", OleDbType.WChar)
            lobjParameterTELEFONO.Value = pstrTELEFONO
            lobjCommand.Parameters.Add(lobjParameterTELEFONO)
            ' Nombre completo
            Dim lobjParameterNOMBRECOMPLETO As New OleDbParameter("@NOMBRECOMPLETO", OleDbType.WChar)
            lobjParameterNOMBRECOMPLETO.Value = pstrNOMBRECOMPLETO
            lobjCommand.Parameters.Add(lobjParameterNOMBRECOMPLETO)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertHCLocal(), mobjCnnSqlServer)

            ' ID
            Dim lobjParameterID As New SqlParameter("@ID", SqlDbType.VarChar)
            lobjParameterID.Value = pstrID
            lobjCommand.Parameters.Add(lobjParameterID)
            ' Historias Clínicas
            Dim lobjParameterHC As New SqlParameter("@HC", SqlDbType.VarChar)
            lobjParameterHC.Value = pstrHC
            lobjCommand.Parameters.Add(lobjParameterHC)
            ' Apellidos
            Dim lobjParameterAPELLIDOS As New SqlParameter("@APELLIDOS", SqlDbType.VarChar)
            lobjParameterAPELLIDOS.Value = pstrAPELLIDOS
            lobjCommand.Parameters.Add(lobjParameterAPELLIDOS)
            ' Nombre
            Dim lobjParameterNOMBRE As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNOMBRE.Value = pstrNOMBRE
            lobjCommand.Parameters.Add(lobjParameterNOMBRE)
            ' Numero SS
            Dim lobjParameterNUMEROSS As New SqlParameter("@NUMEROSS", SqlDbType.VarChar)
            lobjParameterNUMEROSS.Value = pstrNUMEROSS
            lobjCommand.Parameters.Add(lobjParameterNUMEROSS)
            ' Código sexo
            Dim lobjParameterCODIGOSEXO As New SqlParameter("@CODIGOSEXO", SqlDbType.Float)
            lobjParameterCODIGOSEXO.Value = pdblCODIGOSEXO
            lobjCommand.Parameters.Add(lobjParameterCODIGOSEXO)
            ' Fecha Nacimiento
            Dim lobjParameterFECHANAC As New SqlParameter("@FECHANAC", SqlDbType.VarChar)
            lobjParameterFECHANAC.Value = pstrFECHANAC
            lobjCommand.Parameters.Add(lobjParameterFECHANAC)
            ' DNI
            Dim lobjParameterDNI As New SqlParameter("@DNI", SqlDbType.VarChar)
            lobjParameterDNI.Value = pstrDNI
            lobjCommand.Parameters.Add(lobjParameterDNI)
            ' Dirección
            Dim lobjParameterDIRECCION As New SqlParameter("@DIRECCION", SqlDbType.VarChar)
            lobjParameterDIRECCION.Value = pstrDIRECCION
            lobjCommand.Parameters.Add(lobjParameterDIRECCION)
            ' Población
            Dim lobjParameterPOBLACION As New SqlParameter("@POBLACION", SqlDbType.VarChar)
            lobjParameterPOBLACION.Value = pstrPOBLACION
            lobjCommand.Parameters.Add(lobjParameterPOBLACION)
            ' Provincia
            Dim lobjParameterPROVINCIA As New SqlParameter("@PROVINCIA", SqlDbType.VarChar)
            lobjParameterPROVINCIA.Value = pstrPROVINCIA
            lobjCommand.Parameters.Add(lobjParameterPROVINCIA)
            ' Código postal
            Dim lobjParameterCP As New SqlParameter("@CP", SqlDbType.VarChar)
            lobjParameterCP.Value = pstrCP
            lobjCommand.Parameters.Add(lobjParameterCP)
            ' Teléfono
            Dim lobjParameterTELEFONO As New SqlParameter("@TELEFONO", SqlDbType.VarChar)
            lobjParameterTELEFONO.Value = pstrTELEFONO
            lobjCommand.Parameters.Add(lobjParameterTELEFONO)
            ' Nombre completo
            Dim lobjParameterNOMBRECOMPLETO As New SqlParameter("@NOMBRECOMPLETO", SqlDbType.VarChar)
            lobjParameterNOMBRECOMPLETO.Value = pstrNOMBRECOMPLETO
            lobjCommand.Parameters.Add(lobjParameterNOMBRECOMPLETO)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarPerfilMicroLocal
    ' Desc: Rutina que inserta perfil micro a la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Sub InsertarPerfilMicroLocal(ByVal pstrCodigo As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertPerfilMicroLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.WChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertPerfilMicroLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.VarChar)
            lobjParameterCodigo.Value = pstrCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarMicroMuestraLocal
    ' Desc: Rutina que inserta micro-muestra a la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Sub InsertarMicroMuestraLocal(ByVal pdblCodigoPrueba As Double, ByVal pdblCodigoMuestra As Double)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertMicroMuestrasLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO_PRUEBA", OleDbType.Double)
            lobjParameterCodigo.Value = pdblCodigoPrueba
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New OleDbParameter("@CODIGO_MUESTRA", OleDbType.Double)
            lobjParameterAbrv.Value = pdblCodigoMuestra
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertMicroMuestrasLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO_PRUEBA", SqlDbType.Float)
            lobjParameterCodigo.Value = pdblCodigoPrueba
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New SqlParameter("@CODIGO_MUESTRA", SqlDbType.Float)
            lobjParameterAbrv.Value = pdblCodigoMuestra
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' InsertarPruebasMicroLocal
    ' Desc: Rutina que inserta pruebas micro a la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Sub InsertarPruebasMicroLocal(ByVal pdblCodigo As Double, ByVal pstrAbrv As String, ByVal pstrNombre As String)

        If Me.mobjConfig.TipoLocal = 0 Then
            ' Access            
            Dim lobjCommand As New OleDbCommand(sqlInsertPruebasMicroLocal(), mobjCnnAccess)

            Dim lobjParameterCodigo As New OleDbParameter("@CODIGO", OleDbType.Double)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New OleDbParameter("@ABRV", OleDbType.WChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New OleDbParameter("@NOMBRE", OleDbType.WChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        Else
            ' SQL Server
            Dim lobjCommand As New SqlCommand(sqlInsertPruebasMicroLocal(), mobjCnnSqlServer)

            Dim lobjParameterCodigo As New SqlParameter("@CODIGO", SqlDbType.Float)
            lobjParameterCodigo.Value = pdblCodigo
            lobjCommand.Parameters.Add(lobjParameterCodigo)

            Dim lobjParameterAbrv As New SqlParameter("@ABRV", SqlDbType.VarChar)
            lobjParameterAbrv.Value = pstrAbrv
            lobjCommand.Parameters.Add(lobjParameterAbrv)

            Dim lobjParameterNom As New SqlParameter("@NOMBRE", SqlDbType.VarChar)
            lobjParameterNom.Value = pstrNombre
            lobjCommand.Parameters.Add(lobjParameterNom)

            lobjCommand.ExecuteNonQuery()

        End If

    End Sub

    ' ********************************************************************************
    ' BorrarTablaBioquimicaLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaBioquimicaLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeletePruebasBioquimicaLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeletePruebasBioquimicaLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de bioquímica." + ex.Message, "BorrarTablaBioquimicaLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaMuestraLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaMuestraLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteMuestrasLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteMuestrasLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de muestras." + ex.Message, "BorrarTablaMuestraLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaPerfilBioquimicaLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaPerfilBioquimicaLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeletePerfilBioquimicaLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeletePerfilBioquimicaLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de perfiles de bioquímica." + ex.Message, "BorrarTablaPerfilBioquimicaLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaPerfilMicroLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaPerfilMicroLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeletePerfilMicroLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeletePerfilMicroLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de perfiles de microbiología." + ex.Message, "BorrarTablaPerfilMicroLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaMedicosLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaMedicosLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteMedicosLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteMedicosLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de médicos." + ex.Message, "BorrarTablaMedicosLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaDestinosLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaDestinosLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteDestinosLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteDestinosLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de destinos." + ex.Message, "BorrarTablaDestinosLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaOrigenesLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaOrigenesLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteOrigenesLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteOrigenesLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de orígenes." + ex.Message, "BorrarTablaOrigenesLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaCorrelacionesLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaCorrelacionesLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteCorrelacionesLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteCorrelacionesLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de correlaciones." + ex.Message, "BorrarTablaCorrelacionesLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaServiciosLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaServiciosLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteServiciosLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteServiciosLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de servicios." + ex.Message, "BorrarTablaServiciosLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaHCLocal
    ' Desc: Rutina que borra la BBDD local
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaHCLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteHCLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteHCLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de historias clínicas." + ex.Message, "BorrarTablaHCLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaMicroMuestraLocal
    ' Desc: Rutina que borra la BBDD local de micro-muestra
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaMicroMuestraLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeleteMicroMuestrasLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeleteMicroMuestrasLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de micro-muestras." + ex.Message, "BorrarTablaMicroMuestraLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' BorrarTablaPruebasMicroLocal
    ' Desc: Rutina que borra la BBDD local de pruebas de microbiología
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function BorrarTablaPruebasMicroLocal() As Boolean

        Try
            If Me.mobjConfig.TipoLocal = 0 Then
                ' Access                
                Dim lobjCommand As New OleDbCommand(sqlDeletePruebasMicroLocal(), mobjCnnAccess)
                lobjCommand.ExecuteNonQuery()
            Else
                ' Server
                Dim lobjCommand As New SqlCommand(sqlDeletePruebasMicroLocal(), mobjCnnSqlServer)
                lobjCommand.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al borrar la tabla de pruebas micro." + ex.Message, "BorrarTablaPruebasMicroLocal", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End Try

    End Function

    ' ********************************************************************************
    ' sqlSelectPruebasBioquimicaRemotas
    ' Desc: Función que devuelve el SELECT de selección de pruebas de bioquímica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlSelectPruebasBioquimicaRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Bioquimica
            lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3}", .CODIGO, .ABRV, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectMuestrasRemotas
    ' Desc: Función que devuelve el SELECT de selección de muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlSelectMuestrasRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Muestra
            lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3}", .CODIGO, .ABRV, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectPruebasMicroRemotas
    ' Desc: Función que devuelve el SELECT de selección de pruebas de microbiología
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlSelectPruebasMicroRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Microbiologia
            lstrSQL = String.Format("SELECT {0}, {1}, {2} FROM {3}", .CODIGO, .ABRV, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectPerfilBioquimicaRemotas
    ' Desc: Función que devuelve el SELECT de selección de perfiles de bioquímica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlSelectPerfilBioquimicaRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.PerfilBioquimica
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectPerfilMicroRemotas
    ' Desc: Función que devuelve el SELECT de selección de perfiles de microbiología
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectPerfilMicroRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.PerfilMicro
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectMedicosRemotas
    ' Desc: Función que devuelve el SELECT de selección de médicos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectMedicosRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Medicos
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectOrigenesRemotas
    ' Desc: Función que devuelve el SELECT de selección de origenes
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectOrigenesRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Origenes
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectCorrelacionesRemotas
    ' Desc: Función que devuelve el SELECT de selección de correlaciones
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectCorrelacionesRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Correlaciones
            lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10} FROM {11}", _
                                                    .ID, .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, _
                                                    .CODIGO_ORIGEN, .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, .CODIGO_GRUPO_FACTURACION, _
                                                    .CODIGO_CARGO, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectServiciosRemotas
    ' Desc: Función que devuelve el SELECT de selección de servicios
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectServiciosRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Servicios
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectDestinosRemotas
    ' Desc: Función que devuelve el SELECT de selección de destinos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectDestinosRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Destinos
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectDiagnosticosRemotas
    ' Desc: Función que devuelve el SELECT de selección de diagnósticos
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Private Function sqlSelectDiagnosticosRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.Diagnostico
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO, .NOMBRE, .TABLA)
        End With

        Return lstrSQL

    End Function


    ' ********************************************************************************
    ' sqlSelectHCRemotas
    ' Desc: Función que devuelve el SELECT de selección de historias clínicas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlSelectHCRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.HistoriasClinicas
            lstrSQL = String.Format("SELECT {0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13} FROM {14}", _
                                                    .ID, .NUM_HISTORIA, .APELLIDOS, .NOMBRE, .NUM_SS, .COD_SEXO, .FECHA_NACIMIENTO, _
                                                    .DNI, .DIRECCION, .POBLACION, .COD_PROVINCIA, .COD_POSTAL, .TELEFONO, .NOMBRE_COMPLETO, .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertPruebasBioquimicaLocal
    ' Desc: Función que devuelve el INSERT local de pruebas de bioquímica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlInsertPruebasBioquimicaLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Bioquimica
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}, {3}) VALUES (@CODIGO, @ABRV, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .ABRV, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertMuestrasLocal
    ' Desc: Función que devuelve el INSERT local de muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlInsertMuestrasLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Muestra
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}, {3}) VALUES (@CODIGO, @ABRV, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .ABRV, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertPerfilBioquimicaLocal
    ' Desc: Función que devuelve el INSERT local de perfiles de bioquimica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlInsertPerfilBioquimicaLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.PerfilBioquimica
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertPerfilMicroLocal
    ' Desc: Función que devuelve el INSERT local de perfiles de micro
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertPerfilMicroLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.PerfilMicro
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertMedicosLocal
    ' Desc: Función que devuelve el INSERT local de médicos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertMedicosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Medicos
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertOrigenesLocal
    ' Desc: Función que devuelve el INSERT local de origenes
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertOrigenesLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Origenes
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertCorrelacionesLocal
    ' Desc: Función que devuelve el INSERT local de correlaciones
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertCorrelacionesLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Correlaciones
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10},{11}) VALUES " + _
                                                "(@ID, @CAMPO, @DESENCADENANTE, @DOCTOR, @SERVICIO, @ORIGEN, " + _
                                                "@DESTINO, @MOTIVO, @TIPO, @GRUPOFAC, @CARGO)", .TABLA, .ID, _
                                                .CODIGO_CAMPO, .CODIGO_DESENCADENANTE, .CODIGO_DOCTOR, .CODIGO_SERVICIO, _
                                                .CODIGO_ORIGEN, .CODIGO_DESTINO, .CODIGO_MOTIVO, .CODIGO_TIPO, _
                                                .CODIGO_GRUPO_FACTURACION, .CODIGO_CARGO)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertDestinosLocal
    ' Desc: Función que devuelve el INSERT local de destinos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertDestinosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Destinos
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertDiagnosticoLocal
    ' Desc: Función que devuelve el INSERT local de diagnóstico
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Private Function sqlInsertDiagnosticoLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Diagnostico
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function


    ' ********************************************************************************
    ' sqlInsertServiciosLocal
    ' Desc: Función que devuelve el INSERT local de servicios
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertServiciosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Servicios
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertHCLocal
    ' Desc: Función que devuelve el INSERT local de historias clínicas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlInsertHCLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.HistoriasClinicas
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13}, {14}) " + _
                                                 "VALUES (@ID, @HC, @APELLIDOS, @NOMBRE, @NUMEROSS, @CODIGOSEXO, " + _
                                                 "@FECHANAC, @DNI, @DIRECCION, @POBLACION, @PROVINCIA, @CP, @TELEFONO, " + _
                                                 "@NOMBRECOMPLETO)", _
                                                 .TABLA, .ID, .NUM_HISTORIA, .APELLIDOS, .NOMBRE, .NUM_SS, .COD_SEXO, .FECHA_NACIMIENTO, _
                                                 .DNI, .DIRECCION, .POBLACION, .COD_PROVINCIA, .COD_POSTAL, .TELEFONO, .NOMBRE_COMPLETO)
        End With

        Return lstrSQL

    End Function


    ' ********************************************************************************
    ' sqlInsertMicroMuestrasLocal
    ' Desc: Función que devuelve el INSERT local de micro-muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlInsertMicroMuestrasLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.MicroMuestra
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}) VALUES (@CODIGO_PRUEBA, @CODIGO_MUESTRA)", _
                                                    .TABLA, .CODIGO_PRUEBA, .CODIGO_MUESTRA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlInsertPruebasMicroLocal
    ' Desc: Función que devuelve el INSERT local de pruebas de micro
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlInsertPruebasMicroLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Microbiologia
            lstrSQL = String.Format("INSERT INTO {0} ({1}, {2}, {3}) VALUES (@CODIGO, @ABRV, @NOMBRE)", _
                                                    .TABLA, .CODIGO, .ABRV, .NOMBRE)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeletePruebasBioquimicaLocal
    ' Desc: Función que devuelve el DELETE local de pruebas de bioquímica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlDeletePruebasBioquimicaLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Bioquimica
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteMuestrasLocal
    ' Desc: Función que devuelve el DELETE local de muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteMuestrasLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Muestra
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeletePerfilBioquimicaLocal
    ' Desc: Función que devuelve el DELETE local de perfiles de bioquímica
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlDeletePerfilBioquimicaLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.PerfilBioquimica
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeletePerfilMicroLocal
    ' Desc: Función que devuelve el DELETE local de perfiles de micro
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeletePerfilMicroLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.PerfilMicro
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteMedicosLocal
    ' Desc: Función que devuelve el DELETE local de médicos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteMedicosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Medicos
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteOrigenesLocal
    ' Desc: Función que devuelve el DELETE local de orígenes
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteOrigenesLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Origenes
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function


    ' ********************************************************************************
    ' sqlDeleteCorrelacionesLocal
    ' Desc: Función que devuelve el DELETE local de correlaciones
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteCorrelacionesLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Correlaciones
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteServiciosLocal
    ' Desc: Función que devuelve el DELETE local de servicios
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteServiciosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Servicios
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteDestinosLocal
    ' Desc: Función que devuelve el DELETE local de destinos
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteDestinosLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Destinos
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteHCLocal
    ' Desc: Función que devuelve el DELETE local de historias clínicas
    ' NBL: 7/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteHCLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.HistoriasClinicas
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeleteMicroMuestrasLocal
    ' Desc: Función que devuelve el DELETE local de micro-muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlDeleteMicroMuestrasLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.MicroMuestra
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlDeletePruebasMicroLocal
    ' Desc: Función que devuelve el DELETE local de pruebas de micro
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlDeletePruebasMicroLocal() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDLocal.Microbiologia
            lstrSQL = String.Format("DELETE FROM {0}", .TABLA)
        End With

        Return lstrSQL

    End Function

    ' ********************************************************************************
    ' sqlSelectMicroMuestrasRemotas
    ' Desc: Función que devuelve el SELECT de selección de relacion de micro - muestras
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Private Function sqlSelectMicroMuestrasRemotas() As String

        Dim lstrSQL As String

        With Me.mobjConfigBBDDRemota.MicroMuestra
            lstrSQL = String.Format("SELECT {0}, {1} FROM {2}", .CODIGO_PRUEBA, .CODIGO_MUESTRA, .TABLA)
        End With

        Return lstrSQL

    End Function

End Class
