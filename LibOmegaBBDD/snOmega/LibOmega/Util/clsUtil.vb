Imports System.Reflection.Assembly
Imports System.Xml
Imports System.IO
Imports System.Windows.Forms

Public Class clsUtil

    ' ******************************************************************
    ' NoCaseSensitive 
    ' Desc:
    ' NBL 23/6/2008
    ' ******************************************************************
    Public Shared Function NoCaseSensitive() As Integer

        Dim lobjINI As New clsINI

        Dim lintResultado As Integer = CType(lobjINI.IniGet(DLLPath(True) & "LibOmega.ini", "General", "NoCaseSensitive", "0"), Integer)

        Return lintResultado

    End Function

    ' **************************************************************************************************************************
    ' DLLPath(Optional ByVal backSlash As Boolean = False) As String
    ' Descripcion: Función que devuelve la ruta física de la carpeta contenedora de la dll que
    '                            se está ejecutando.
    ' Parámetros: backSlash: Se utiliza para devolver la ruta con un "\" final en el caso que sea True
    ' Autor: NBL
    ' Fecha de creación: 25/5/2006
    ' **************************************************************************************************************************
    Public Shared Function DLLPath(Optional ByVal backSlash As Boolean = False) As String

        Dim s As String = IO.Path.GetDirectoryName(GetExecutingAssembly.GetName.CodeBase.ToString)

        If s.StartsWith("file") Then s = s.Substring(6)
        ' si hay que añadirle el backslash
        If backSlash Then
            s &= "\"
        End If

        Return s

    End Function

    Public Shared Function Right(ByVal value As String, ByVal length As Integer) As String
        If (length < 0) Then
            Throw New ArgumentException("Length is too short.")
        End If
        If (length = 0 OrElse value Is Nothing) Then
            Return ""
        End If
        Dim size As Integer = value.Length
        If (length >= size) Then
            Return value
        End If
        Return value.Substring(size - length, length)
    End Function

    ' ************************************************************************
    ' CargarConfiguracion
    ' Desc: Rutina que carga la configuración de BBDD
    ' NBL: 11/01/2007
    ' ************************************************************************
    Public Shared Sub CargarConfiguracion(ByRef mobjConfig As clsConfig, ByVal mstrNombreArchivoConfig As String)

        If Not My.Computer.FileSystem.FileExists(mstrNombreArchivoConfig) Then
            MessageBox.Show("El archivo de configuración no existe" + vbCrLf + "Entre en la configuración antes de ejectuar los diálogos", _
                                        "Omega", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Try
            Dim xmlReader As New XmlTextReader(mstrNombreArchivoConfig)
            Dim Reader As New Serialization.XmlSerializer(GetType(clsConfig))
            mobjConfig = Reader.Deserialize(xmlReader)
            xmlReader.Close()
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al cargar la configuración" + vbCrLf + "Entre en la configuración de la librería", _
                                        "Omega", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub

    ' ************************************************************************
    ' CargarConfiguracion
    ' Desc: Rutina que carga la configuración de BBDD
    ' NBL: 11/01/2007
    ' ************************************************************************
    Public Shared Sub CargarConfiguracionBBDD(ByRef pobjConfigBBDD As clsConfigBBDD, ByVal pstrNombreArchivoConfigBBDD As String)

        If Not My.Computer.FileSystem.FileExists(pstrNombreArchivoConfigBBDD) Then
            MessageBox.Show("El archivo de configuración de BBDD no existe" + vbCrLf + "Entre en la configuración de BBDD antes de ejectuar los diálogos", _
                                        "Omega", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If

        Try
            'Leer un archivo XML y cargarlo en un objeto
            Dim xmlReader As New XmlTextReader(pstrNombreArchivoConfigBBDD)
            'Crear un objeto para deserializar el archivo XML
            Dim Reader As New Serialization.XmlSerializer(GetType(clsConfigBBDD))
            'Deserialziar el archivo xml y cargarlo en un objeto
            pobjConfigBBDD = Reader.Deserialize(xmlReader)
            'Cerrar Archivo XML
            xmlReader.Close()
        Catch ex As Exception
            MessageBox.Show("Ha habido un error al cargar la configuración de BBDD" + vbCrLf + "Entre en la configuración de la librería", _
                                        "Omega", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Try

    End Sub

End Class
