'<ComClass(ComClsSincroniza.ClassId, ComClsSincroniza.InterfaceId, ComClsSincroniza.EventsId)> _
Public Class ComClsSincroniza

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "258b5a2f-e645-4a82-82e6-8104d572ff31"
    Public Const InterfaceId As String = "acc4f1d1-3a3e-4da1-b983-f7ca912deed0"
    Public Const EventsId As String = "2423a610-8b51-47e7-a899-e8d82d9ed3c3"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()

        MyBase.New()

        mobjSincroniza = New clsSincroniza

    End Sub

    ' Defino el objeto que va a ser el que sincronice
    Private mobjSincroniza As clsSincroniza

    ' ********************************************************************************
    ' SincronizaPruebasBioquimica
    ' Desc: Función que sincroniza las tablas locales de pruebas de bioquimica con las remotas
    ' NBL: 5/3/2007
    ' ********************************************************************************
    Public Function SincronizaPruebasBioquimica() As Boolean

        Return mobjSincroniza.SincronizaPruebasBioquimica()

    End Function

    ' ********************************************************************************
    ' SincronizaMuestras
    ' Desc: Función que sincroniza las tablas locales de muestras con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaMuestras() As Boolean

        Return mobjSincroniza.SincronizaMuestras()

    End Function

    ' ********************************************************************************
    ' SincronizaPruebasMicro
    ' Desc: Función que sincroniza las tablas locales de pruebas de micro con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaPruebasMicro() As Boolean

        Return mobjSincroniza.SincronizaPruebasMicro()

    End Function

    ' ********************************************************************************
    ' SincronizaMicroMuestra
    ' Desc: Función que sincroniza las tablas locales de relación de micro-muestra con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaMicroMuestra() As Boolean

        Return mobjSincroniza.SincronizaMicroMuestras()

    End Function

    ' ********************************************************************************
    ' SincronizaPerfilesBio
    ' Desc: Función que sincroniza las tablas locales de perfiles de bioquímica con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaPerfilesBio() As Boolean

        Return mobjSincroniza.SincronizaPerfilBioquimica()

    End Function

    ' ********************************************************************************
    ' SincronizaPerfilesMicro
    ' Desc: Función que sincroniza las tablas locales de perfiles de micro con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaPerfilesMicro() As Boolean

        Return mobjSincroniza.SincronizaPerfilMicro()

    End Function

    ' ********************************************************************************
    ' SincronizaMedicos
    ' Desc: Función que sincroniza las tablas locales de medicos con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaMedicos() As Boolean

        Return mobjSincroniza.SincronizaMedicos()

    End Function

    ' ********************************************************************************
    ' SincronizaHC
    ' Desc: Función que sincroniza las tablas locales de historias clínicas con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaHC() As Boolean

        Return mobjSincroniza.SincronizaHC()

    End Function

    ' ********************************************************************************
    ' SincronizaCorrelaciones
    ' Desc: Función que sincroniza las tablas locales de correlaciones con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaCorrelaciones() As Boolean

        Return mobjSincroniza.SincronizaCorrelaciones()

    End Function

    ' ********************************************************************************
    ' SincronizaOrigenes
    ' Desc: Función que sincroniza las tablas locales de orígenes con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaOrigenes() As Boolean

        Return mobjSincroniza.SincronizaOrigenes()

    End Function

    ' ********************************************************************************
    ' SincronizaServicios
    ' Desc: Función que sincroniza las tablas locales de servicios con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaServicios() As Boolean

        Return mobjSincroniza.SincronizaServicios()

    End Function

    ' ********************************************************************************
    ' SincronizaDestinos
    ' Desc: Función que sincroniza las tablas locales de destinos con las remotas
    ' NBL: 12/3/2007
    ' ********************************************************************************
    Public Function SincronizaDestinos() As Boolean

        Return mobjSincroniza.SincronizaDestinos()

    End Function

    ' ********************************************************************************
    ' SincronizaDiagnosticos
    ' Desc: Función que sincroniza las tablas locales de diagnosticos con las remotas
    ' NBL: 12/6/2007
    ' ********************************************************************************
    Public Function SincronizaDiagnosticos() As Boolean

        Return mobjSincroniza.SincronizaDiagnosticos()

    End Function

End Class


