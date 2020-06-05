Public Class clsSincronizaEventsArgs

    Inherits EventArgs

    Private _NumeroRegistro As Integer = 0
    Private _NumeroTotalRegistros As Integer = 0
    Private _TablaSincro As TablaSincro

    Public Sub New(ByVal pintNumeroRegistro As Integer, ByVal pintNumeroTotalRegistros As Integer, ByVal penTablaSincro As TablaSincro)

        Me._NumeroRegistro = pintNumeroRegistro
        Me._NumeroTotalRegistros = pintNumeroTotalRegistros
        Me._TablaSincro = penTablaSincro

    End Sub

    ' Ahora creamos las propiedades del evento
    Public ReadOnly Property NumeroRegistro() As Integer
        Get
            Return _NumeroRegistro
        End Get
    End Property

    Public ReadOnly Property NumeroTotalRegistros() As Integer
        Get
            Return _NumeroTotalRegistros
        End Get
    End Property

    Public ReadOnly Property TablaSincroniza() As TablaSincro
        Get
            Return _TablaSincro
        End Get
    End Property

End Class
