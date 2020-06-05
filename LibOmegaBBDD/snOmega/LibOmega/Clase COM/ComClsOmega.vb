Imports System.Windows.Forms
'<ComClass(ComClsSincroniza.ClassId, ComClsSincroniza.InterfaceId, ComClsSincroniza.EventsId)> _
Public Class ComClsOmega

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "85e81619-c056-4855-b4c5-e5182e3598b3"
    Public Const InterfaceId As String = "b9ebc03a-624c-4844-9738-698a467beb56"
    Public Const EventsId As String = "8822b8c4-aaa6-443c-91e4-bf5dc8ba74e1"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    ' ************************************************************************
    ' ConsultaMedico
    ' Desc: Función que consulta un médico
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function ConsultaMedico(ByVal pstrCodigoMedico As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaMedico_1(pstrCodigoMedico)

    End Function

    ' ************************************************************************
    ' ConsultaDiagnostico
    ' Desc: Función que consulta un diagnóstico
    ' NBL: 13/06/2007
    ' ************************************************************************
    Public Function ConsultaDiagnostico(ByVal pstrCodigoDiagnostico As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaDiagnostico_1(pstrCodigoDiagnostico)

    End Function

    ' ************************************************************************
    ' ConsultaServicio
    ' Desc: Función que consulta un servicio
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function ConsultaServicio(ByVal pstrCodigoServicio As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaServicio_1(pstrCodigoServicio)

    End Function

    ' ************************************************************************
    ' ConsultaOrigen
    ' Desc: Función que consulta un origen
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function ConsultaOrigen(ByVal pstrCodigoOrigen As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaOrigen_1(pstrCodigoOrigen)

    End Function

    ' ************************************************************************
    ' ConsultaDestino
    ' Desc: Función que muestra el diálogo de destinos
    ' NBL: 21/05/2007
    ' ************************************************************************
    Public Function ConsultaDestino(ByVal pstrCodigoDestino As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaDestino_1(pstrCodigoDestino)

    End Function

    ' ************************************************************************
    ' ConsultaNHUSAbyNHC
    ' Desc: Función que devuelve NHUSA a partir de NHC 
    ' NBL: 16/11/2009
    ' ************************************************************************
    Public Function ConsultaNHUSAbyNHC(ByVal pstrNHC As String, ByVal pstrRutaINI As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.getNHUSAbyNHC(pstrNHC, pstrRutaINI)

    End Function

    ' ************************************************************************
    ' ConsultaNHCbyNHUSA
    ' Desc: Función que devuelve NHC a partir de NHUSA
    ' NBL: 16/11/2009
    ' ************************************************************************
    Public Function ConsultaNHCbyNHUSA(ByVal pstrNHUSA As String, ByVal pstrRutaINI As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.getNHCbyNHUSA(pstrNHUSA, pstrRutaINI)

    End Function

    ' ************************************************************************
    ' ConsultaCorrelacion
    ' Desc: Función que busca una correlación
    ' NBL: 21/07/2009
    ' ************************************************************************
    Public Function ConsultaCorrelacion(ByVal pstrCodigoCorrelacion As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaCorrelacion(pstrCodigoCorrelacion)

    End Function

    ' ************************************************************************
    ' CapturaMedicos
    ' Desc: Función que muestra el diálogo de médicos
    ' NBL: 9/01/2007
    ' ************************************************************************
    Public Function CapturaMedicos(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                        ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación 
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As New ArrayList
        Dim lobjFormulario As DialogoMedicos

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaMedico(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            lobjFormulario = New DialogoMedicos(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return "*"
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' CapturaCorrelacion
    ' Desc: Función que busca una correlación y la devuelve los campos del registro 
    ' NBL: 2/03/2007
    ' ************************************************************************
    Public Function CapturaCorrelacion(ByVal pstrTextoBusqueda As String) As String

        Dim lobjConsulta As New clsConsulta
        Return lobjConsulta.BuscaCorrelacion(pstrTextoBusqueda.Trim)

    End Function

    ' ************************************************************************
    ' CapturaDestinos
    ' Desc: Función que muestra el diálogo de destinos
    ' NBL: 1/03/2007
    ' ************************************************************************
    Public Function CapturaDestinos(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                        ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación 
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As New ArrayList
        Dim lobjFormulario As DialogoDestinos

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaDestino(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            lobjFormulario = New DialogoDestinos(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return "*"
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' CapturaHC
    ' Desc: Función que muestra el diálogo de historias clínicas
    ' NBL: 2/03/2007
    ' ************************************************************************
    Public Function CapturaHC(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                        ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación 
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As ArrayList = Nothing
        Dim lobjFormulario As DialogoHistoriaClinica

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaHC(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            lobjFormulario = New DialogoHistoriaClinica(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Dim Resultado As String = lobjListViewItem.Tag
                Return Resultado
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return "*"
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' CapturaServicios
    ' Desc: Función que muestra el diálogo de Servicios
    ' NBL: 1/03/2007
    ' ************************************************************************
    Public Function CapturaServicios(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                        ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación 
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As New ArrayList
        Dim lobjFormulario As DialogoServicios

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaServicio(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            lobjFormulario = New DialogoServicios(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return "*"
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' CapturaDiagnosticos
    ' Desc: Función que muestra el diálogo de diagnósticos
    ' NBL: 1/03/2007
    ' ************************************************************************
    Public Function CapturaDiagnosticos(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                            ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As ArrayList = Nothing
        Dim lobjFormulario As DialogoDiagnostico

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaDiagnostico(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            lobjFormulario = New DialogoDiagnostico(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjlistviewitem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjlistviewitem.Text + "|" + lobjlistviewitem.SubItems(1).Text
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return ""
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' CapturaOrigenes
    ' Desc: Función que muestra el diálogo de Origenes
    ' NBL: 1/03/2007
    ' ************************************************************************
    Public Function CapturaOrigenes(ByVal pstrTituloDialogo As String, ByVal pstrLabel As String, _
                                                        ByVal pstrPreBusqueda As String, ByVal pbolMuestraDialogo As Boolean) As String

        ' Hacemos una pequeña validación 
        If pstrPreBusqueda.Length = 0 And Not pbolMuestraDialogo Then Return ""

        Dim lobjConsulta As New clsConsulta
        Dim larlstResultado As New ArrayList
        Dim lobjFormulario As DialogoOrigenes

        If pstrPreBusqueda.Length > 0 Then larlstResultado = lobjConsulta.BuscaOrigen(pstrPreBusqueda.Trim)

        If pbolMuestraDialogo Then
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            lobjFormulario = New DialogoOrigenes(pstrTituloDialogo, pstrLabel, pstrPreBusqueda, larlstResultado)
            If lobjFormulario.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Return lobjFormulario.Resultado
            End If
        Else
            If larlstResultado Is Nothing Then Return ""
            If larlstResultado.Count = 0 Then Return ""
            If larlstResultado.Count = 1 Then
                ' Devolvemos lo que hemos obtenido
                Dim lobjListViewItem As ListViewItem = CType(larlstResultado.Item(0), ListViewItem)
                Return lobjListViewItem.Text + "|" + lobjListViewItem.SubItems(1).Text
            End If
            If larlstResultado.Count > 1 Then
                ' Más de un resultado sin mostrar diálogo
                Return "*"
            End If
        End If

        Return ""

    End Function

    ' ************************************************************************
    ' DialogoPruebas
    ' Desc: Función que muestra el diálogo de pruebas
    ' NBL: 9/01/2007
    ' ************************************************************************
    Public Function DialogoPruebas(ByVal pstrTituloDialogo As String, ByVal pstrPruebasPreseleccionadas As String) As String

        Dim lobjDialogoPruebas As New DialogoPruebas(pstrTituloDialogo, pstrPruebasPreseleccionadas)

        Select Case lobjDialogoPruebas.ShowDialog()

            Case Windows.Forms.DialogResult.OK
                Return lobjDialogoPruebas.Resultado

        End Select

        Return ""

    End Function

    ' ************************************************************************
    ' DialogoConfiguracion
    ' Desc: Rutina que muestra el diálogo de configuración de los diálogos
    ' NBL: 16/01/2007
    ' ************************************************************************
    Public Sub DialogoConfiguracion()

        Dim lobjDialogoConfiguracion As New DialogoConfiguracion()
        lobjDialogoConfiguracion.ShowDialog()

    End Sub

    ' ************************************************************************
    ' DialogoConfiguracionBBDD
    ' Desc: Rutina que muestra el diálogo de configuración de nombres de la BBDD
    ' NBL: 9/01/2007
    ' ************************************************************************
    Public Sub DialogoConfiguracionBBDD(ByVal pintLocal As Integer)

        If pintLocal = 0 Then
            Dim lobjDialogoConfiguracionBBDD As New DialogoConfiguracionBBDD(BBDD.Remota)
            lobjDialogoConfiguracionBBDD.ShowDialog()
        Else
            Dim lobjDialogoConfiguracionBBDD As New DialogoConfiguracionBBDD(BBDD.Local)
            lobjDialogoConfiguracionBBDD.ShowDialog()
        End If

    End Sub
End Class


