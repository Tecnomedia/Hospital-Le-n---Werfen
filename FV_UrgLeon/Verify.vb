Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports LibFlexibarNETObjects

Public Class Verify

    Implements LibFlexibarNETObjects.IFlexiValidator

    ' ***************************************************************************************
    ' DECLARACIÓN DE VARIABLES
    ' ***************************************************************************************

    ' Variable donde guardaremos el batch
    Dim mobjFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch
    ' Variable donde guardamos el objeto de aplicación
    Dim mobjFlexibarApp As LibFlexibarNETObjects.FlexibarApp
    Dim mintIndex As Integer
    Dim mobjViewer As System.Drawing.Size
    Dim mobjImageSize As System.Drawing.Size
    '    Dim mobjBBDD As F_BBDD.Consultas
    Dim mobjTraductor As F_Util.MarkToCode
    Dim mstrPlantilla As String

    ' Instancia a la clase de reordenación de columnas del listview manuscrito
    Private lvwColumnSorter1 As ListViewColumnSorter
    ' Instancia a la clase de reordenación de columnas del listview marcas
    Private lvwColumnSorter2 As ListViewColumnSorter

    ' NBL 26/6/2009 Pongo este control para la carga rápida de pruebas
    'Dim mobjOmega As New LibOmega.clsConsulta
    'Dim mobjOmega2 As New LibOmega.ComClsOmega
    Dim mobjConsultasUtil As New FS_Shared.FS_Util
    ' NBL 30/09/2010
    Dim mobjUtilLabs As New F_Util.Labs

    Dim mbolMaximized As Boolean = False

    ' NBL 27/09/2011 
    Dim mobjUtilMarks As New F_Util.MarkToCode

    Dim mobjOmega As New LibOmega.clsConsulta
    Dim mobjOmega2 As New LibOmega.ComClsOmega


#Region "Eventos y funciones IFlexiValidator"

    Public Event BestFit(ByVal sender As Object, ByVal e As System.EventArgs) Implements LibFlexibarNETObjects.IFlexiValidator.BestFit

    Public Event ChangeMark(ByVal sender As Object, ByVal e As LibFlexibarNETObjects.MarkChangedArguments) Implements LibFlexibarNETObjects.IFlexiValidator.ChangeMark

    Public Event ChangeVerifyMode(ByVal sender As Object, ByVal e As LibFlexibarNETObjects.VerifyModeArguments) Implements LibFlexibarNETObjects.IFlexiValidator.ChangeVerifyMode

    Public Sub Load_UserControl(ByRef pobjFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch,
                                ByRef pobjFlexibarApp As LibFlexibarNETObjects.FlexibarApp,
                                ByVal pintIndex As Integer,
                                ByVal pViewerSize As System.Drawing.Size,
                                ByVal pImageSize As System.Drawing.Size,
                                ByRef pobjFlexibarGlobal As FlexibarGlobal) Implements LibFlexibarNETObjects.IFlexiValidator.Load_UserControl

        mobjFlexibarApp = pobjFlexibarApp
        InicializarImagen(pobjFlexibarBatch, pintIndex, pViewerSize, pImageSize)

    End Sub

    Public Function NavigationControl(ByVal pintIndex As Integer, ByVal pintNextIndex As Integer) As Boolean Implements LibFlexibarNETObjects.IFlexiValidator.NavigationControl

        Return True

    End Function

    Public Event PrintImage(ByVal sender As Object, ByVal e As LibFlexibarNETObjects.PrintArguments) Implements LibFlexibarNETObjects.IFlexiValidator.PrintImage

    Public Event SetZoom(ByVal sender As Object, ByVal e As LibFlexibarNETObjects.ZoomArguments) Implements LibFlexibarNETObjects.IFlexiValidator.SetZoom
    Public Event VerifyNavigation As IFlexiValidator.VerifyNavigationEventHandler Implements IFlexiValidator.VerifyNavigation
    Public Event VerifyRefreshCounters As IFlexiValidator.VerifyRefreshCountersEventHandler Implements IFlexiValidator.VerifyRefreshCounters
    Public Event TransferBatch As IFlexiValidator.TransferBatchEventHandler Implements IFlexiValidator.TransferBatch
    Public Event HighLightBlocks As IFlexiValidator.HighLightBlocksEventHandler Implements IFlexiValidator.HighLightBlocks

    Public Sub ViewerMarkChanged(ByVal pstrName As String, ByVal pbolMarked As Boolean) Implements LibFlexibarNETObjects.IFlexiValidator.ViewerMarkChanged

        If Regex.IsMatch(pstrName, "[M,P]\d{4}[M,P,L]") Then
            CambiaSeleccionMarcas(pstrName, pbolMarked)
        End If

        ColorearVerificacion(Me.mobjFlexibarBatch.Images(Me.mintIndex - 1))

    End Sub

#End Region

    ' ************************************************************************************************
    ' CambiaSeleccionMarcas
    ' Desc: Cambiamos la selección de las marcas por el visor
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub CambiaSeleccionMarcas(ByVal pstrNombreMarca As String, ByVal pbolMarked As Boolean)

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        If lobjImage.ARData.TemplateName = "LE_URGENCIAS" Then

            If Regex.IsMatch(pstrNombreMarca, "^P0\d{3}P_") Then
                mobjUtilMarks.CalculaPruebasHemato(lobjImage)
                mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcasHemato"),
                                                   lobjImage.VirtualFields.GetFieldValue("ACodigosMarcasHematoDesc"), Me.lvPruebasMarcasHematologia, lobjImage)
            ElseIf Regex.IsMatch(pstrNombreMarca, "^[M,P]1\d{3}[M,P]_") Then
                ' Calculamos las pruebas de bioquímica y micro de omega
                lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", mobjUtilMarks.MarkTraductor(lobjImage.ARData.ARCheckmarkFields, lobjImage.ARData.TemplateName, False, 1, ""))
                lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), True))
                mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"),
                                                                            lobjImage.VirtualFields.GetFieldValue("ACodigosMarcasDesc"),
                                                                            Me.lvPruebasMarcas, lobjImage)
            End If

        End If

        Me.tabVerificador.SelectedIndex = 1

        'If lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA" Or lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA_2" Then
        '    ' Aquí el tratamiento de las pruebas es diferentes
        '    Dim lstrPruebas As String = ""
        '    Dim lstrObs As String = ""
        '    mobjConsultasUtil.CalculaPruebasMicroVirgenPuerto(lobjImage, lstrPruebas, lstrObs)
        '    ' NBL 17/11/2010 Si es urocultivo hay que añadir una prueba de bioquímica
        '    If lstrPruebas.Contains(",22^M|76|") Then
        '        lstrPruebas &= ",9055^B||"
        '    End If
        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", lstrPruebas)
        '    lobjImage.VirtualFields.SetFieldValue("DObservaciones2", lstrObs)
        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), True))
        '    mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), lobjImage.VirtualFields.GetFieldValue("ACodigosMarcasDesc"), Me.lvPruebasMarcas, lobjImage)

        '    ' Código para liberar las demás marcas en el caso de que esté marcada un tipo de puestra o una localización
        '    ' ESTE CODIGO NO LO PUEDO IMPLEMENTAR POR BUG EN EL FLEXIBAR
        '    'If Regex.IsMatch(pstrNombreMarca, "^[P]\d{4}[M]") And pbolMarked Then
        '    '    For Each lobjMark As LibFlexibarNETObjects.ARCheckmarkField In lobjImage.ARData.ARCheckmarkFields
        '    '        If (lobjMark.Name <> pstrNombreMarca) And Regex.IsMatch(lobjMark.Name, "^[P]\d{4}[M]") Then
        '    '            RaiseEvent ChangeMark(Me, New LibFlexibarNETObjects.MarkChangedArguments(lobjMark.Name, False))
        '    '        End If
        '    '    Next
        '    'ElseIf Regex.IsMatch(pstrNombreMarca, "^[P]\d{4}[L]") And pbolMarked Then
        '    '    For Each lobjMark As LibFlexibarNETObjects.ARCheckmarkField In lobjImage.ARData.ARCheckmarkFields
        '    '        If (lobjMark.Name <> pstrNombreMarca) And Regex.IsMatch(lobjMark.Name, "^[P]\d{4}[L]") Then
        '    '            RaiseEvent ChangeMark(Me, New LibFlexibarNETObjects.MarkChangedArguments(lobjMark.Name, False))
        '    '        End If
        '    '    Next
        '    'End If

        'Else

        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", mobjTraductor.MarkTraductor(lobjImage.ARData.ARCheckmarkFields, lobjImage.ARData.TemplateName, 1))
        '    'CargarCodigosListView(Me.lvPruebasMarcas, lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), 1)
        '    'CargarPruebasSeleccionadas(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), Me.lvPruebasMarcas)

        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), False))
        '    mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), lobjImage.VirtualFields.GetFieldValue("ACodigosMarcasDesc"), Me.lvPruebasMarcas, lobjImage)
        'End If

        'Me.tabVerificador.SelectedIndex = 1

    End Sub

    ' ***********************************************************************************************
    ' InicializarImagen
    ' Desc: Llamada cuando iniciamos la visualización de datos de una nueva imagen
    ' NBL 27/8/2010
    ' ***********************************************************************************************
    Private Sub InicializarImagen(ByRef pobjFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch,
                                                    ByVal pintIndex As Integer, ByVal pViewerSize As System.Drawing.Size,
                                                    ByVal pImageSize As System.Drawing.Size)

        ' En primer lugar almacenamos el árbol de batch y su índice
        Me.mobjFlexibarBatch = pobjFlexibarBatch
        Me.mintIndex = pintIndex
        Me.mobjViewer = pViewerSize
        Me.mobjImageSize = pImageSize

        InicializarCamposVerificador()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(pintIndex - 1)

        Me.mstrPlantilla = ""
        If Not lobjImage.ARData Is Nothing Then
            If Not lobjImage.ARData.TemplateName Is Nothing Then
                Me.mstrPlantilla = lobjImage.ARData.TemplateName
                Me.lblTemplateName.Text = mstrPlantilla

                'If lobjImage.ARData.ARCheckmarkFields.GetFieldValue("P1060P_ADA") = "1" Then
                '    MessageBox.Show("Marcada ADA en liquidos biológicos", "Especializada", MessageBoxButtons.OK, MessageBoxIcon.Information)
                'End If

            End If
        End If

        '' Cargamos los datos de la petición
        CargarDatosPeticion()
        ColorearVerificacion(lobjImage)

        ' NBL 7/12/2009 La imagen ya ha sido revisada, por lo que la ponemos como revisada
        lobjImage.VirtualFields.SetFieldValue("XRevisado", "1")

        ' NBL 29/07/2010 Revisar los demográficos
        If lobjImage.VirtualFields.GetFieldValue("XDemograficos") = "1" Then
            lobjImage.VirtualFields.SetFieldValue("XDemograficos", "0")
            MessageBox.Show("Revise los demográficos", "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Sub

    ' ************************************************************************************************
    ' ColorearVerificacion
    ' Desc: Coloreamos los campos que toquen 
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub ColorearVerificacion(ByRef pobjImage As LibFlexibarNETObjects.Image)

        If Regex.IsMatch(pobjImage.VirtualFields.GetFieldValue("DNoPet"), "^\d{8}$") Then
            Me.txtNumeroPeticion.BackColor = Color.PaleGreen
        Else
            Me.txtNumeroPeticion.BackColor = Color.LightSalmon
        End If

        If pobjImage.VirtualFields.GetFieldValue("DNoHist").Trim.Length > 0 Then
            Me.txtNombre.BackColor = Color.PaleGreen
            Me.txtApellido1.BackColor = Color.PaleGreen
            Me.txtSexo.BackColor = Color.PaleGreen
            Me.txtFechaNacimiento.BackColor = Color.PaleGreen
            Me.txtNHC1.BackColor = Color.PaleGreen
            Me.txtNSS.BackColor = Color.PaleGreen
            Me.txtDNI.BackColor = Color.PaleGreen
            'Me.txtNHC2.BackColor = Color.PaleGreen
            'Me.txtNHC3.BackColor = Color.PaleGreen
            'Me.txtICU.BackColor = Color.PaleGreen
        Else
            Me.txtNombre.BackColor = Color.Moccasin
            Me.txtApellido1.BackColor = Color.Moccasin
            Me.txtSexo.BackColor = Color.Moccasin
            Me.txtFechaNacimiento.BackColor = Color.Moccasin
            Me.txtNHC1.BackColor = Color.Moccasin
            Me.txtNSS.BackColor = Color.Moccasin
            Me.txtDNI.BackColor = Color.Moccasin
            'Me.txtNHC2.BackColor = Color.Moccasin
            'Me.txtNHC3.BackColor = Color.Moccasin
            'Me.txtICU.BackColor = Color.Moccasin
        End If

        ' DOCTOR
        If pobjImage.VirtualFields.GetFieldValue("DDoctor").Trim.Length > 0 Then
            Me.txtCodigoCustom1.BackColor = Color.PaleGreen
            Me.txtNombreCustom1.BackColor = Color.PaleGreen
        Else
            Me.txtCodigoCustom1.BackColor = Color.Moccasin
            Me.txtNombreCustom1.BackColor = Color.Moccasin
        End If

        ' DESTINO
        If pobjImage.VirtualFields.GetFieldValue("DDestino").Trim.Length > 0 Then
            Me.txtCodigoCustom2.BackColor = Color.PaleGreen
            Me.txtNombreCustom2.BackColor = Color.PaleGreen
        Else
            Me.txtCodigoCustom2.BackColor = Color.Moccasin
            Me.txtNombreCustom2.BackColor = Color.Moccasin
        End If

        ' ORIGEN
        If pobjImage.VirtualFields.GetFieldValue("DOrigen").Trim.Length > 0 Then
            Me.txtCodigoCustom3.BackColor = Color.PaleGreen
            Me.txtNombreCustom3.BackColor = Color.PaleGreen
        Else
            Me.txtCodigoCustom3.BackColor = Color.Moccasin
            Me.txtNombreCustom3.BackColor = Color.Moccasin
        End If

        ' SERVICIO
        If pobjImage.VirtualFields.GetFieldValue("DServicio").Trim.Length > 0 Then
            Me.txtServicio.BackColor = Color.PaleGreen
            Me.txtNombreCustom4.BackColor = Color.PaleGreen
        Else
            Me.txtServicio.BackColor = Color.Moccasin
            Me.txtNombreCustom4.BackColor = Color.Moccasin
        End If

        ' EPISODIO
        If pobjImage.VirtualFields.GetFieldValue("DEpisodio").Trim.Length > 0 Then
            Me.txtCodigoCustom5.BackColor = Color.PaleGreen
        Else
            Me.txtCodigoCustom5.BackColor = Color.Moccasin
        End If

        ' CAMA
        If pobjImage.VirtualFields.GetFieldValue("DCama").Trim.Length > 0 Then
            Me.txtCama.BackColor = Color.PaleGreen
        Else
            Me.txtCama.BackColor = Color.Moccasin
        End If

        Me.tabVerificador.Refresh()

    End Sub

    ' ************************************************************************************************
    ' CargarDatosPeticion
    ' Desc: Mostramos los datos de la petición en el interface de verificación
    ' Para los tres tabs
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosPeticion()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        CargarDatosTab1(lobjImage)
        CargarDatosTab2(lobjImage)
        CargarDatosTab3(lobjImage)

    End Sub

    ' ************************************************************************************************
    ' CargarDatosTab3
    ' Desc: Cargamos los datos de la pestaña 3 que contiene los datos del paciente y de la petición
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosTab3(ByRef pobjImage As LibFlexibarNETObjects.Image)

        'Me.txtNumeroPeticion3.Text = pobjImage.VirtualFields.GetFieldValue("DNoPet")

        '' Ponemos la orina 24 en el caso de que sea el form 67
        'Me.txtOrina24.Text = pobjImage.VirtualFields.GetFieldValue("AOrina24")
        Me.txtObservaciones.Text = pobjImage.VirtualFields.GetFieldValue("DObservaciones")
        Me.txtDiagnosticos.Text = pobjImage.VirtualFields.GetFieldValue("DCDiagnostico")

        'Me.txtSemanaGestacion.Text = pobjImage.VirtualFields.GetFieldValue("D_SemanaGestacion")
        'Me.txtVolumen.Text = pobjImage.VirtualFields.GetFieldValue("D_VOLUMEN_ORINA")

    End Sub

    ' ************************************************************************************************
    ' CargarPruebasSeleccionadas
    ' Desc: Rutina que carga las pruebas seleccionadas en el listview de pruebas de marcas
    ' NBL 20/10/2009
    ' ************************************************************************************************
    Private Sub CargarPruebasSeleccionadas(ByVal pstrListadoCodigos As String, ByVal pobjListView As ListView)

        Dim lstrLista As New ArrayList()
        pobjListView.Items.Clear()
        ' Primero miramos si hay pruebas o no
        If pstrListadoCodigos.Trim.Length = 0 Then Exit Sub
        ' Hemos de hacer un split de comas para separar pruebas
        Dim lstrPruebas() As String = pstrListadoCodigos.Split(",")
        If lstrPruebas.Length <= 1 Then Exit Sub
        Dim lstrCodigoPrueba As String = "", lstrTipoPrueba As String = "", lstrCodigoMuestra As String = ""
        Dim lstrAbrvPrueba As String = "", lstrDescripcionPrueba As String = "", lstrAbrvMuestra As String = "", lstrDescripcionMuestra As String = ""

        For lintContador As Integer = 1 To lstrPruebas.Length - 1
            Dim lstrPrueba() As String = lstrPruebas(lintContador).Split("|")
            Dim lstrCodigoP() As String = lstrPrueba(0).Split("^")
            lstrCodigoPrueba = lstrCodigoP(0)
            lstrTipoPrueba = lstrCodigoP(1)
            lstrCodigoMuestra = lstrPrueba(1)
            If lstrCodigoPrueba.Trim.Length <> 0 Then
                Dim lobjListViewItem As New ListViewItem()
                lobjListViewItem.Text = lstrCodigoPrueba

                Dim lbolMicro As Boolean = False
                Dim lbolPerfil As Boolean = False

                If lstrTipoPrueba = "M" Then
                    lbolMicro = True
                Else
                    lbolMicro = False
                End If

                If lstrCodigoMuestra.Trim.Length = 0 Then
                    lbolPerfil = True
                Else
                    lbolPerfil = False
                End If

                'mobjOmega.getAbrvDescripcionPruebas(lstrCodigoPrueba, lstrAbrvPrueba, lstrDescripcionPrueba, _
                '                                                                lstrCodigoMuestra, lstrAbrvMuestra, lstrDescripcionMuestra, "", lbolMicro, lbolPerfil)
                lobjListViewItem.SubItems.Add(lstrAbrvPrueba)
                lobjListViewItem.SubItems.Add(lstrDescripcionPrueba)
                lobjListViewItem.SubItems.Add(lstrCodigoMuestra)
                lobjListViewItem.SubItems.Add(lstrAbrvMuestra)
                lobjListViewItem.SubItems.Add(lstrDescripcionMuestra)
                lstrLista.Add(lobjListViewItem)
            End If
        Next

        With pobjListView
            .BeginUpdate()
            .SuspendLayout()
            .Items.AddRange(lstrLista.ToArray(GetType(ListViewItem)))
            .EndUpdate()
            .ResumeLayout()
        End With

    End Sub

    ' ************************************************************************************************
    ' CargarDatosTab2
    ' Desc: Cargamos los datos de la pestaña 1 que contiene los datos del paciente y de la petición
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosTab2(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Me.txtNumeroPeticion2.Text = pobjImage.VirtualFields.GetFieldValue("DNoPet")
        'CargarPruebasSeleccionadas(pobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), Me.lvPruebasMarcas)
        'CargarPruebasSeleccionadas(pobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), Me.lvPruebasManuscritas)
        mobjUtilLabs.CargarPruebasListView(pobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), pobjImage.VirtualFields.GetFieldValue("ACodigosMarcasDesc"), Me.lvPruebasMarcas, pobjImage)
        mobjUtilLabs.CargarPruebasListView(pobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), pobjImage.VirtualFields.GetFieldValue("ACodigosManuscritoDesc"), Me.lvPruebasManuscritas, pobjImage)
        'CargarCodigosListView(Me.lvPruebasManuscritas, pobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), 2)
        'CargarCodigosListView(Me.lvPruebasMarcas, pobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), 1)

        ' NBL 26/09/2011 ------------------------------------------------------------------------------------------------------
        mobjUtilLabs.CargarPruebasListView(pobjImage.VirtualFields.GetFieldValue("ACodigosMarcasHemato"),
                                           pobjImage.VirtualFields.GetFieldValue("ACodigosMarcasHematoDesc"), Me.lvPruebasMarcasHematologia, pobjImage)

        ' -------------------------------------------------------------------------------------------------------------------------

    End Sub

    ' ************************************************************************************************
    ' CargarDatosTab1
    ' Desc: Cargamos los datos de la pestaña 1 que contiene los datos del paciente y de la petición
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosTab1(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Me.txtNumeroPeticion.Text = pobjImage.VirtualFields.GetFieldValue("DNoPet")

        CargarDatosPaciente(pobjImage)

        CargarDatosMedicoTab1(pobjImage)
        CargarDatosServicioTab1(pobjImage)
        CargarDatosDestinoTab1(pobjImage)
        'CargarDatosOrigenTab1(pobjImage)
        ' NBL 19/01/2010 Cargo el episodio
        'Me.txtCodigoCustom5.Text = pobjImage.VirtualFields.GetFieldValue("DEpisodio")
        'Me.txtCodigoCustom6.Text = pobjImage.VirtualFields.GetFieldValue("DCama")

        ' NBL 8/11/2010 Cargo la cama
        Me.txtCama.Text = pobjImage.VirtualFields.GetFieldValue("DCama")

        'Me.txtNHCCentro.Text = pobjImage.VirtualFields.GetFieldValue("DNoHist2").Trim

    End Sub

    ' ************************************************************************************************
    ' CargarDatosPacienteTab1
    ' Desc: Cargamos los datos del paciente
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosPaciente(ByRef pobjImage As LibFlexibarNETObjects.Image)

        '' NBL Si se han solicitado los datos por webservice, no dejamos hacer nada en el diálogo de verificación
        'If pobjImage.VirtualFields.GetFieldValue("webservice") = "1" Then
        '    Me.txtBuscar.ReadOnly = True
        '    Me.txtBuscar.Text = "Webservice"
        'Else
        '    Me.txtBuscar.ReadOnly = False
        '    Me.txtBuscar.Text = ""
        'End If

        ''Me.txtBuscar.Text = ""
        If pobjImage.VirtualFields.GetFieldValue("DNoHist").Trim.Length > 0 Then
            Me.txtNombre.Text = pobjImage.VirtualFields.GetFieldValue("DNombre")
            Me.txtApellido1.Text = pobjImage.VirtualFields.GetFieldValue("DApellido1").Trim & " " & pobjImage.VirtualFields.GetFieldValue("DApellido2").Trim
            Me.txtSexo.Text = pobjImage.VirtualFields.GetFieldValue("DSexo")
            Dim lstrFechaNac As String = pobjImage.VirtualFields.GetFieldValue("DFechaNac")
            If lstrFechaNac.Length = 8 Then
                Me.txtFechaNacimiento.Text = lstrFechaNac.Substring(6, 2) & "/" & lstrFechaNac.Substring(4, 2) & "/" & lstrFechaNac.Substring(0, 4)
            ElseIf lstrFechaNac.Length = 10 Then
                Me.txtFechaNacimiento.Text = lstrFechaNac
            Else
                Me.txtFechaNacimiento.Text = ""
            End If
            Me.txtNHC1.Text = pobjImage.VirtualFields.GetFieldValue("DNoHist")
            Me.txtNSS.Text = pobjImage.VirtualFields.GetFieldValue("DNoSS")
            Me.txtDNI.Text = pobjImage.VirtualFields.GetFieldValue("DDNI")
            'Me.txtNHC2.Text = pobjImage.VirtualFields.GetFieldValue("DNoHist2")
        Else
            Me.txtNombre.Text = ""
            Me.txtApellido1.Text = ""
            Me.txtSexo.Text = ""
            Me.txtFechaNacimiento.Text = ""
            Me.txtNHC1.Text = ""
            Me.txtNSS.Text = ""
            Me.txtDNI.Text = ""
            'Me.txtNHC2.Text = pobjImage.VirtualFields.GetFieldValue("DNoHist2")
        End If

    End Sub

    ' ************************************************************************************************
    ' CargarDatosDestinoTab1
    ' Desc: Cargamos los datos del destino
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosDestinoTab1(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrDestino As String = pobjImage.VirtualFields.GetFieldValue("DDestino")

        If lstrDestino.Trim.Length > 0 Then
            Me.txtCodigoCustom2.Text = lstrDestino
            Me.txtNombreCustom2.Text = pobjImage.VirtualFields.GetFieldValue("TDestino")
            Me.txtBuscarCustom2.Text = ""
        Else
            If Not pobjImage.ARData.ARTextFields Is Nothing Then
                Me.txtBuscarCustom2.Text = pobjImage.ARData.ARTextFields.GetFieldValue("ADestino")
            End If
            Me.txtCodigoCustom2.Text = ""
            Me.txtNombreCustom2.Text = ""
        End If

    End Sub

    ' ************************************************************************************************
    ' CargarDatosOrigenTab1
    ' Desc: Cargamos los datos del origen
    ' NBL 16/7/2009
    ' ************************************************************************************************
    Private Sub CargarDatosOrigenTab1(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrOrigen As String = pobjImage.VirtualFields.GetFieldValue("DOrigen")

        If lstrOrigen.Trim.Length > 0 Then
            Me.txtCodigoCustom3.Text = lstrOrigen
            Me.txtNombreCustom3.Text = pobjImage.VirtualFields.GetFieldValue("TOrigen")
            Me.txtBuscarCustom3.Text = ""
        Else
            If Not pobjImage.ARData.ARTextFields Is Nothing Then
                Me.txtBuscarCustom3.Text = pobjImage.ARData.ARTextFields.GetFieldValue("AOrigen")
            End If
            Me.txtCodigoCustom3.Text = ""
            Me.txtNombreCustom3.Text = ""
        End If

    End Sub

    ' ************************************************************************************************
    ' CargarDatosServicioTab1
    ' Desc: Cargamos los datos del origen
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosServicioTab1(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrServicio As String = pobjImage.VirtualFields.GetFieldValue("DServicio")

        If lstrServicio.Trim.Length > 0 Then
            Me.txtServicio.Text = lstrServicio
            'Me.txtNombreCustom4.Text = pobjImage.VirtualFields.GetFieldValue("TServicio")
            'Me.txtBuscarCustom4.Text = ""
        Else
            'If Not pobjImage.ARData.ARTextFields Is Nothing Then
            'Me.txtBuscarCustom4.Text = pobjImage.ARData.ARTextFields.GetFieldValue("AServicio")
            'End If
            Me.txtServicio.Text = ""
            'Me.txtNombreCustom4.Text = ""
        End If

    End Sub

    ' ************************************************************************************************
    ' CargarDatosMedicoTab1
    ' Desc: Cargamos los datos del médico
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargarDatosMedicoTab1(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrDoctor As String = pobjImage.VirtualFields.GetFieldValue("DDoctor")

        If lstrDoctor.Trim.Length > 0 Then
            Me.txtCodigoCustom1.Text = lstrDoctor
            Me.txtNombreCustom1.Text = pobjImage.VirtualFields.GetFieldValue("DTDoctor")
            Me.txtBuscarCustom1.Text = ""
        Else
            If Not pobjImage.ARData.ARTextFields Is Nothing Then
                Me.txtBuscarCustom1.Text = pobjImage.ARData.ARTextFields.GetFieldValue("ADoctor")
            End If
            Me.txtCodigoCustom1.Text = ""
            Me.txtNombreCustom1.Text = ""
        End If

    End Sub

    Private Sub lvPruebasManuscritas_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvPruebasManuscritas.ColumnClick

        ' Indicamos que no es una fecha
        lvwColumnSorter1.OrderDate = False

        ' Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter1.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter1.Order = SortOrder.Ascending) Then
                lvwColumnSorter1.Order = SortOrder.Descending
            Else
                lvwColumnSorter1.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter1.SortColumn = e.Column
            lvwColumnSorter1.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.lvPruebasManuscritas.Sort()

    End Sub

    Private Sub lvPruebasMarcas_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvPruebasMarcas.ColumnClick
        ' Indicamos que no es una fecha
        lvwColumnSorter2.OrderDate = False

        ' Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter2.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter2.Order = SortOrder.Ascending) Then
                lvwColumnSorter2.Order = SortOrder.Descending
            Else
                lvwColumnSorter2.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter2.SortColumn = e.Column
            lvwColumnSorter2.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.lvPruebasMarcas.Sort()
    End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        InicializarCamposVerificador()

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter1 = New ListViewColumnSorter()
        Me.lvPruebasManuscritas.ListViewItemSorter = lvwColumnSorter1

        ' Mostramos los datos ordenados de los apellidos
        lvwColumnSorter1.SortColumn = 0
        lvwColumnSorter1.Order = SortOrder.Ascending

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter2 = New ListViewColumnSorter()
        Me.lvPruebasMarcas.ListViewItemSorter = lvwColumnSorter2

        ' Mostramos los datos ordenados de los apellidos
        lvwColumnSorter2.SortColumn = 0
        lvwColumnSorter2.Order = SortOrder.Ascending

        ' mobjTraductor = New F_Util.MarkToCode()

        'mobjOmega = New LibOmega.clsConsulta
        '  mobjOmega.IniciarCargaPruebas()

    End Sub

    ' ***********************************************************************************************
    ' InicializarCamposVerificador
    ' Desc: Ponemos el texto o deshabilitamos los campos que no nos interesen
    ' NBL 8/5/2009
    ' ***********************************************************************************************
    Private Sub InicializarCamposVerificador()

        Me.txtBuscar.Text = ""
        Me.txtNHC1.Text = ""
        Me.txtNombre.Text = ""
        Me.txtApellido1.Text = ""
        Me.txtSexo.Text = ""
        Me.txtFechaNacimiento.Text = ""
        Me.txtNSS.Text = ""
        Me.txtDNI.Text = ""
        Me.txtCodigoCustom1.Text = ""
        Me.txtNombreCustom1.Text = ""
        Me.txtCodigoCustom2.Text = ""
        Me.txtNombreCustom2.Text = ""
        Me.txtServicio.Text = ""
        Me.txtCama.Text = ""
        Me.txtDiagnosticos.Text = ""
        Me.txtObservaciones.Text = ""

    End Sub

    Protected Overrides Sub Finalize()

        'mobjOmega.CerrarCargaPruebasgaPruebas()
        MyBase.Finalize()

    End Sub

    Private Sub btnModoVerificacion_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        RaiseEvent ChangeVerifyMode(Me, New LibFlexibarNETObjects.VerifyModeArguments(LibFlexibarNETObjects.enVerifyMode.Maximized))

    End Sub

    ' ***********************************************************************************************
    ' CampoModificado
    ' Desc: 
    ' NBL 13/5/2009
    ' ***********************************************************************************************
    Private Sub CampoModificado(ByVal pstrNombreCampo As String, ByVal pstrValor As String)

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue(pstrNombreCampo, pstrValor)
        ColorearVerificacion(lobjImage)

    End Sub

    ' ***********************************************************************************************
    ' NumeroPeticionModificado
    ' Desc: Rutina que se llama cuando se modifica el número de petición
    ' NBL 13/5/2009
    ' ***********************************************************************************************
    Private Sub NumeroPeticionModificado(ByVal pstrValorModificacion As String)

        Me.txtNumeroPeticion2.Text = pstrValorModificacion
        'Me.txtNumeroPeticion3.Text = pstrValorModificacion
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DNoPet", pstrValorModificacion)
        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub txtNumeroPeticion_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub txtNumeroPeticion_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtNumeroPeticion.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.txtBuscar.Focus()
        End If

    End Sub

    Private Sub txtNumeroPeticion_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtNumeroPeticion.Validating
        NumeroPeticionModificado(Me.txtNumeroPeticion.Text.Trim)
    End Sub

    ' *************************************************************************************************
    ' ResetMedico
    ' Desc: 
    ' NBL 18/5/2009
    ' *************************************************************************************************
    Private Sub ResetMedico()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DDoctor", "")
        lobjImage.VirtualFields.SetFieldValue("DTDoctor", "")
        CargarDatosMedicoTab1(lobjImage)
        ColorearVerificacion(lobjImage)

    End Sub

    ' *************************************************************************************************
    ' ResetServicio
    ' Desc: 
    ' NBL 18/5/2009
    ' *************************************************************************************************
    Private Sub ResetServicio()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DServicio", "")
        lobjImage.VirtualFields.SetFieldValue("TServicio", "")
        CargarDatosServicioTab1(lobjImage)
        ColorearVerificacion(lobjImage)

    End Sub

    ' *************************************************************************************************
    ' ResetDestino
    ' Desc: 
    ' NBL 18/5/2009
    ' *************************************************************************************************
    Private Sub ResetDestino()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DDestino", "")
        lobjImage.VirtualFields.SetFieldValue("TDestino", "")
        CargarDatosDestinoTab1(lobjImage)
        ColorearVerificacion(lobjImage)

    End Sub

    ' *************************************************************************************************
    ' ResetOrigen
    ' Desc: 
    ' NBL 16/7/2009
    ' *************************************************************************************************
    Private Sub ResetOrigen()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DOrigen", "")
        lobjImage.VirtualFields.SetFieldValue("TOrigen", "")
        CargarDatosOrigenTab1(lobjImage)
        ColorearVerificacion(lobjImage)

    End Sub

    ' *************************************************************************************************
    ' ResetDemograficos
    ' Desc: 
    ' NBL 18/5/2009
    ' *************************************************************************************************
    Private Sub ResetDemograficos()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("DNombre", "")
        lobjImage.VirtualFields.SetFieldValue("DApellido1", "")
        lobjImage.VirtualFields.SetFieldValue("DApellido2", "")
        lobjImage.VirtualFields.SetFieldValue("DSexo", "")
        lobjImage.VirtualFields.SetFieldValue("DFechaNac", "")
        lobjImage.VirtualFields.SetFieldValue("DNoHist", "")
        lobjImage.VirtualFields.SetFieldValue("DNoSS", "")
        lobjImage.VirtualFields.SetFieldValue("DNoHist2", "")
        lobjImage.VirtualFields.SetFieldValue("DNoHist3", "")
        lobjImage.VirtualFields.SetFieldValue("DDNI", "")
        CargarDatosPaciente(lobjImage)
        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub btnResetDemograficos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetDemograficos.Click
        ResetDemograficos()
    End Sub

    'Private Sub txtPeso_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then Me.txtTalla.Focus()
    'End Sub

    ' PVT: Ejemplo de validación de formato decimal 
    Private Sub txtPeso_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Not e.KeyChar.IsDigit(e.KeyChar) And Not e.KeyChar = Convert.ToChar(8) Then

            e.Handled = UtilGlobal.UShared.HandledDecimalFormat(sender, e, True, True)

            'If (e.KeyChar = Convert.ToChar(".") Or e.KeyChar = Convert.ToChar(",")) AndAlso (txtPeso.Text.IndexOf(".") = -1 And txtPeso.Text.IndexOf(",") = -1) Then
            '    e.Handled = False
            'Else
            '    e.Handled = True
            'End If
        End If
    End Sub

    'Private Sub btnTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTransfer.Click
    '    Dim pintIndex As Integer = mintIndex - 1
    '    Dim blnLogin As Boolean = False
    '    Dim strUser As String = ""

    '    Dim objShared As New FS_Shared.FS_Util

    '    'If mstrActivate = "1" Then
    '    '    Dim objUser As New FS_Cruces.UserAuthentication(mstrDSN, mstrTable, mstrUserField, mstrPasswordField, CInt(mstrTimeOut))
    '    '    If Not objUser.checkUser(CInt(mstrMinutesRefresh)) Then
    '    '        Exit Sub
    '    '    Else
    '    '        blnLogin = True

    '    '        strUser = objUser.User
    '    '    End If
    '    'End If

    '    Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

    '    If objShared.ValidarTransferencia(lobjImage, Me.mobjFlexibarBatch, Me.mintIndex - 1) Then

    '        Dim lobjASTM As New ASTMnet.Constructor

    '        Try


    '            If Not lobjImage.RemovePage Then

    '                ' Actualizamos las observaciones y los diagnósticos
    '                Dim strObservaciones As String = UtilGlobal.UShared.ClearSpecialChars(lobjImage.VirtualFields.GetFieldValue("DObservaciones"))
    '                lobjImage.VirtualFields.SetFieldValue("DObservaciones", strObservaciones)


    '                strObservaciones = UtilGlobal.UShared.ClearSpecialChars(lobjImage.VirtualFields.GetFieldValue("DTDiagnostico"))
    '                lobjImage.VirtualFields.SetFieldValue("DTDiagnostico", strObservaciones)


    '                ' En el caso de que sea una petición de especializada:

    '                Dim strPruebas As String = lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas")
    '                strPruebas = "," & strPruebas
    '                lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", "," & strPruebas.Replace(",,", ","))

    '                strPruebas = lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito")
    '                strPruebas = "," & strPruebas
    '                lobjImage.VirtualFields.SetFieldValue("ACodigosManuscrito", "," & strPruebas.Replace(",,", ","))

    '                Dim lstrFechaNac As String = lobjImage.VirtualFields.GetFieldValue("DFechaNac")

    '                If Not String.IsNullOrEmpty(lstrFechaNac) And lstrFechaNac.Length = 10 Then
    '                    lstrFechaNac = lstrFechaNac.Substring(6, 4) & lstrFechaNac.Substring(3, 2) & lstrFechaNac.Substring(0, 2)
    '                    lobjImage.VirtualFields.SetFieldValue("DFechaNac", lstrFechaNac)
    '                End If

    '                '    Dim intParam As Integer = CType(UtilGlobal.UShared.getParameterFromIni("", "Parameters", "ParamNumber"), Integer)

    '                '    For i As Integer = 1 To intParam
    '                '        Dim lstrNombreAux As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_NombreAux", "")
    '                '        Dim lstrCodigoPrueba As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_CodigoPrueba", "")
    '                '        Dim lstrString1PxxxxA As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_String1PxxxxA", "")
    '                '        Dim lstrString2PxxxxA As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_String2PxxxxA", "")
    '                '        Dim lstrString1PxxxxR As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_String1PxxxxR", "")
    '                '        Dim lstrString2PxxxxR As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_String2PxxxxR", "")
    '                '        Dim lstrString3PxxxxR As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_String3PxxxxR", "")
    '                '        ' NBL 18082010 en este caso solo enviamos la linea R
    '                '        Dim lstrSoloR As String = UtilGlobal.UShared.getParameterFromIni("", "Parameters", "Param" & i.ToString & "_SoloR", "")

    '                '        ' Si hay Diuresis voy a crear un PxxxxA y un PxxxxR como campos virtuales para que lo entienda la exportación

    '                '        If lobjImage.VirtualFields.GetFieldValue(lstrNombreAux) <> "" Then

    '                '            Dim lstrValue As String = lobjImage.VirtualFields.GetFieldValue(lstrNombreAux)

    '                '            If Not lobjImage.VirtualFields.FieldExists("P000" & i.ToString & "A") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("P000" & i.ToString & "A", ""))
    '                '            If Not lobjImage.VirtualFields.FieldExists("P0001" & i.ToString & "R") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("P000" & i.ToString & "R", ""))

    '                '            If lstrSoloR = "1" Then
    '                '                lobjImage.VirtualFields.SetFieldValue("P000" & i.ToString & "A", "")
    '                '            Else
    '                '                lobjImage.VirtualFields.SetFieldValue("P000" & i.ToString & "A", lstrString1PxxxxA & lstrCodigoPrueba & lstrString2PxxxxA)
    '                '            End If

    '                '            lobjImage.VirtualFields.SetFieldValue("P000" & i.ToString & "R", lstrString1PxxxxR & lstrCodigoPrueba & lstrString2PxxxxR & lstrValue & lstrString3PxxxxR)
    '                '        End If
    '                '    Next

    '                'End If

    '                'If mstrActivate = "1" And mstrASTM = "1" And blnLogin = True Then
    '                '    lobjImage.VirtualFields.SetFieldValue("DUser", strUser)
    '                'End If

    '                Dim strExportMessage As String = ""

    '                Dim strErrorMessage As String = ""
    '                If Not lobjASTM.ExportarASTM(lobjImage, Me.mintIndex - 1, Me.mobjFlexibarApp, Me.mobjFlexibarBatch, strErrorMessage) Then
    '                    If Not lobjASTM.ExportarASTM(lobjImage, mintIndex - 1, Me.mobjFlexibarApp, Me.mobjFlexibarBatch, strExportMessage) Then
    '                        MsgBox("Se ha producido un error en la transferencia de la petición." & vbCrLf & "Motivo : " & strExportMessage, MsgBoxStyle.Exclamation, "Transferencia de Petición.")
    '                        lobjImage.VirtualFields.SetFieldValue("DUser", "")
    '                    Else
    '                        lobjImage.VirtualFields.SetFieldValue("DExportado", "1")

    '                        'If mstrActivate = "1" And mstrLog = "1" And blnLogin = True Then
    '                        '    FS_Cruces.UserAuthentication.writeUserLog(strUser, lobjImage.VirtualFields.GetFieldValue("DNoPet"))
    '                        'End If

    '                        'Me.InicializarImagen(Me.mobjFlexibarBatch, IIf(pintIndex Mod 2 = 0, pintIndex + 3, pintIndex + 2), Me.mobjViewer, Me.mobjImageSize)
    '                        Me.InicializarImagen(Me.mobjFlexibarBatch, pintIndex + 1, Me.mobjViewer, Me.mobjImageSize)
    '                    End If
    '                Else
    '                    MsgBox("No se ha podido realizar la transferencia de la imágen." & vbCrLf & "Motivo:" & strErrorMessage, MsgBoxStyle.Exclamation, "Transferencia individual")
    '                End If
    '            End If


    '        Catch ex As Exception

    '            MsgBox(ex.Message & vbCrLf & ex.StackTrace)

    '            'pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
    '            'pExportProcessResult.ErrorDescription = ex.Message
    '        End Try

    '    End If
    'End Sub

    Private Sub btnTransfer2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTransfer2.Click
        'Me.btnTransfer_Click(sender, e)
    End Sub

    Private Sub btnModoVerificacion_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModoVerificacion.Click

        CambiarModoVerificacion()

    End Sub

    Private Sub tabVerificador_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles tabVerificador.DrawItem

        'Firstly we'll define some parameters.
        Dim CurrentTab As TabPage = tabVerificador.TabPages(e.Index)
        Dim ItemRect As Rectangle = tabVerificador.GetTabRect(e.Index)

        Dim lobjColor As New Color
        ColorearTabs(e.Index, lobjColor)

        Dim FillBrush As New SolidBrush(lobjColor)
        Dim TextBrush As New SolidBrush(Color.Black)
        Dim sf As New StringFormat
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        'If we are currently painting the Selected TabItem we'll 
        'change the brush colors and inflate the rectangle.
        If CBool(e.State And DrawItemState.Selected) Then
            '    FillBrush.Color = Color.White
            '    TextBrush.Color = Color.Red
            ItemRect.Inflate(2, 2)
        End If

        ''Set up rotation for left and right aligned tabs
        If tabVerificador.Alignment = TabAlignment.Left Or tabVerificador.Alignment = TabAlignment.Right Then
            Dim RotateAngle As Single = 90
            If tabVerificador.Alignment = TabAlignment.Left Then RotateAngle = 270
            Dim cp As New PointF(ItemRect.Left + (ItemRect.Width \ 2), ItemRect.Top + (ItemRect.Height \ 2))
            e.Graphics.TranslateTransform(cp.X, cp.Y)
            e.Graphics.RotateTransform(RotateAngle)
            ItemRect = New Rectangle(-(ItemRect.Height \ 2), -(ItemRect.Width \ 2), ItemRect.Height, ItemRect.Width)
        End If

        'Next we'll paint the TabItem with our Fill Brush
        e.Graphics.FillRectangle(FillBrush, ItemRect)

        'Now draw the text.
        e.Graphics.DrawString(CurrentTab.Text, e.Font, TextBrush, RectangleF.op_Implicit(ItemRect), sf)

        'Reset any Graphics rotation
        e.Graphics.ResetTransform()

        ''Finally, we should Dispose of our brushes.
        FillBrush.Dispose()
        TextBrush.Dispose()

    End Sub

    ' ************************************************************************************************
    ' ColorearTabs
    ' Desc: Coloreamos los tabs
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub ColorearTabs(ByVal pintIndexTab As String, ByRef pobjColor As System.Drawing.Color)

        Select Case pintIndexTab
            Case 0
                ColorearTab1(pobjColor)
            Case 1
                ColorearTab2(pobjColor)
            Case 2
                ColorearTab3(pobjColor)
        End Select

    End Sub

    ' ************************************************************************************************
    ' ColorearTab1
    ' Desc: Coloreamos los tabs
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub ColorearTab1(ByRef pobjColor As System.Drawing.Color)

        Dim lobjColor As New Color

        If Me.txtNumeroPeticion.BackColor = Color.LightSalmon Then
            pobjColor = Color.LightSalmon
            Exit Sub
        End If

        For Each lobjControl As Control In Me.gbDatosPaciente.Controls
            If lobjControl.BackColor = Color.Moccasin Then
                pobjColor = Color.Moccasin
                Exit Sub
            End If
        Next

        For Each lobjControl As Control In Me.gbDatosPeticion.Controls
            If lobjControl.BackColor = Color.Moccasin Then
                pobjColor = Color.Moccasin
                Exit Sub
            End If
        Next

        pobjColor = Color.PaleGreen

    End Sub

    ' ************************************************************************************************
    ' ColorearTab2
    ' Desc: Coloreamos los tabs
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub ColorearTab2(ByRef pobjColor As System.Drawing.Color)

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        pobjColor = Color.PaleGreen

        If Not lobjImage.ARData.ARCheckmarkFields Is Nothing And Not lobjImage.ARData.TemplateName Is Nothing Then

            Dim lobjINI As New UtilGlobal.clsINI

            Dim lintNumeroMarcasNegro As Integer = CInt(lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", lobjImage.ARData.TemplateName, "Numero", 0))

            If lintNumeroMarcasNegro > 0 Then
                For lintContador As Integer = 1 To lintNumeroMarcasNegro
                    If lobjImage.ARData.ARCheckmarkFields.GetFieldValue("QN_" & Microsoft.VisualBasic.Right("00" & lintContador.ToString, 2)) = "1" And lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito").Trim = "" Then
                        pobjColor = Color.Moccasin
                        Exit For
                    End If
                Next
            End If

        End If

    End Sub

    ' ************************************************************************************************
    ' ColorearTab3
    ' Desc: Coloreamos los tabs
    ' NBL 8/5/2009
    ' ************************************************************************************************
    Private Sub ColorearTab3(ByRef pobjColor As System.Drawing.Color)

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        pobjColor = Color.PaleGreen

        If Not lobjImage.ARData.ARCheckmarkFields Is Nothing And Not lobjImage.ARData.TemplateName Is Nothing Then

            Dim lobjINI As New UtilGlobal.clsINI

            Dim lintNumeroMarcasNegro As Integer = CInt(lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", lobjImage.ARData.TemplateName, "Numero", 0))

            If lintNumeroMarcasNegro > 0 Then
                For lintContador As Integer = 1 To lintNumeroMarcasNegro
                    If lobjImage.ARData.ARCheckmarkFields.GetFieldValue("QN_" & Microsoft.VisualBasic.Right("00" & lintContador.ToString, 2)) = "1" Then
                        pobjColor = Color.Moccasin
                        Exit For
                    End If
                Next
            End If

        End If

    End Sub

    Private Sub txtModoVerificacion2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtModoVerificacion2.Click

        CambiarModoVerificacion()

    End Sub

    ' **********************************************************************************************************************
    ' CambiarModoVerificacion
    ' NBL 21/10/2010
    ' Desc.: Cambiamos el modo de verificación
    ' **********************************************************************************************************************
    Private Sub CambiarModoVerificacion()

        If Me.mbolMaximized = False Then
            Me.mbolMaximized = True
            RaiseEvent ChangeVerifyMode(Me, New LibFlexibarNETObjects.VerifyModeArguments(LibFlexibarNETObjects.enVerifyMode.Maximized))
        Else
            Me.mbolMaximized = False
            RaiseEvent ChangeVerifyMode(Me, New LibFlexibarNETObjects.VerifyModeArguments(LibFlexibarNETObjects.enVerifyMode.Normal))
        End If

    End Sub

    Private Sub btnAddPruebaManuscrita_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddPruebaManuscrita.Click

        DialogoPruebas()

    End Sub

    ' ************************************************************************************************
    ' DialogoPruebas
    ' Desc: Rutina que gestiona el diálogo de selección de pruebas
    ' NBL 14/5/2009
    ' ************************************************************************************************
    Private Sub DialogoPruebas()

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        ' Los códigos manuscritos están separados ,6^B||,1010^B||

        Dim lstrResultado As String = mobjOmega2.DialogoPruebas("Selección pruebas manuscritas",
                                                  lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito").ToString())
        lobjImage.VirtualFields.SetFieldValue("ACodigosManuscrito", lstrResultado)
        'CargarCodigosListView(Me.lvPruebasManuscritas, lstrResultado, 2)
        'CargarPruebasSeleccionadas(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito").ToString(), Me.lvPruebasManuscritas)

        If Not lobjImage.ARData Is Nothing AndAlso Not lobjImage.ARData.TemplateName Is Nothing AndAlso lobjImage.ARData.TemplateName.Trim <> "" Then
            If lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA" Or lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA_2" Then
                'CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), True))
            Else
                'CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), False))
            End If
        Else
            'CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), False))
        End If

        CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), False))

        mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), lobjImage.VirtualFields.GetFieldValue("ACodigosManuscritoDesc"), Me.lvPruebasManuscritas, lobjImage)

    End Sub

    ' ***********************************************************************************************
    ' EliminarPruebasManuscritaSeleccionadas
    ' Desc: Eliminamos las pruebas que estén seleccionadas en el listview
    ' NBL 14/5/2009
    ' ***********************************************************************************************
    Private Sub EliminarPruebasManuscritaSeleccionadas()

        If MessageBox.Show("¿Eliminar pruebas seleccionadas?", "Pruebas manuscritas", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then Exit Sub

        If Me.lvPruebasManuscritas.SelectedItems.Count > 0 Then
            For Each lobjListView As ListViewItem In Me.lvPruebasManuscritas.SelectedItems
                lobjListView.Remove()
            Next
        End If

        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(mintIndex - 1)

        Dim lstrINI As String = UtilGlobal.UShared.GetFolderLaboCFG & "\" & lobjImage.ARData.TemplateName & ".ini"

        ' Con las que quedan, hacemos un bucle y las metemos en la variable virtual que las sustenta
        Dim lstrCodigos As String = ""

        If Me.lvPruebasManuscritas.Items.Count > 0 Then
            For Each lobjListView As ListViewItem In Me.lvPruebasManuscritas.Items

                If lobjListView.SubItems.Count > 3 Then
                    lstrCodigos &= "," & lobjListView.Text.Trim & "^" & lobjListView.SubItems(6).Text & "|" & lobjListView.SubItems(3).Text & "|"
                Else
                    lstrCodigos &= "," & lobjListView.Text.Trim
                End If

            Next
            'lstrCodigos = lstrCodigos.Substring(1)
        End If
        CampoModificado("ACodigosManuscrito", lstrCodigos)

        If Not lobjImage.ARData Is Nothing AndAlso Not lobjImage.ARData.TemplateName Is Nothing AndAlso lobjImage.ARData.TemplateName.Trim <> "" Then
            If lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA" Or lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA_2" Then
                '   CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), True))
            Else
                '  CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), False))
            End If
        Else
            'CampoModificado("ACodigosManuscritoDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), False))
        End If

        mobjUtilLabs.CargarPruebasListView(lobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"), lobjImage.VirtualFields.GetFieldValue("ACodigosManuscritoDesc"), Me.lvPruebasManuscritas, lobjImage)

    End Sub

    Private Sub btnEliminarTodasPruebasManuscritas_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminarTodasPruebasManuscritas.Click

        EliminarTodasPruebasManuscrita()

    End Sub

    ' ***********************************************************************************************
    ' EliminarTodasPruebasManuscrita
    ' Desc: Eliminamos las pruebas que estén seleccionadas en el listview
    ' NBL 14/5/2009
    ' ***********************************************************************************************
    Private Sub EliminarTodasPruebasManuscrita()

        If MessageBox.Show("¿Eliminar todas las pruebas seleccionadas?", "Pruebas manuscritas", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then Exit Sub
        Me.lvPruebasManuscritas.Items.Clear()
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(mintIndex - 1)
        lobjImage.VirtualFields.SetFieldValue("ACodigosManuscrito", "")
        lobjImage.VirtualFields.SetFieldValue("ACodigosManuscritoDesc", "")

    End Sub

    Private Sub btnEliminarPruebaManuscrita_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEliminarPruebaManuscrita.Click

        EliminarPruebasManuscritaSeleccionadas()

    End Sub

    Private Sub txtBuscar_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscar.KeyDown

        If Me.txtBuscar.Text.Trim.Length > 0 And e.KeyCode = Keys.Enter Then

            Try
                DialogoPaciente()
            Catch ex As Exception
                MessageBox.Show("Error en la consulta de paciente" & vbCrLf & ex.Source & vbCrLf & ex.Message,
                                                "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try

        End If

    End Sub

    ' ************************************************************************************************
    ' DialogoPaciente
    ' Desc: Rutina que gestiona el diálogo de pacientes
    ' NBL 11/5/2009
    ' ************************************************************************************************
    Private Sub DialogoPaciente()

        Dim mobjShared As New FS_Shared.FS_Util
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        mobjShared.RecuperaDatosPaciente(lobjImage, Me.txtBuscar.Text.Trim)
        '    Dim lstrResultado As String = mobjOmega2.CapturaHC("Búsqueda de pacientes", "Texto de búsqueda", Me.txtBuscar.Text.Trim.ToUpper, True)
        '   If lstrResultado.Trim.Length > 0 Then CargaDatosPacienteConsulta(lstrResultado)

        If lobjImage.VirtualFields.GetFieldValue("DNoHist").Trim.Length > 0 Then
            CargarDatosPaciente(lobjImage)
        End If

        Me.txtBuscar.Text = ""

    End Sub

    ' ************************************************************************************************
    ' CargaDatosPacienteConsulta
    ' Desc: Rutina que carga en los controles del verificador el resultado de la consulta de paciente
    ' NBL 11/5/2009
    ' ************************************************************************************************
    Private Sub CargaDatosPacienteConsulta(ByVal pstrResultado As String)

        Dim lstrDatosPaciente() As String = pstrResultado.Split("|")
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)

        If lstrDatosPaciente.Length > 6 Then

            lobjImage.VirtualFields.SetFieldValue("DNombre", lstrDatosPaciente(3).ToString())
            lobjImage.VirtualFields.SetFieldValue("DApellido1", lstrDatosPaciente(2).ToString())
            lobjImage.VirtualFields.SetFieldValue("DApellido2", lstrDatosPaciente(2).ToString())
            lobjImage.VirtualFields.SetFieldValue("DSexo", lstrDatosPaciente(5).ToString())
            'lobjImage.VirtualFields.SetFieldValue("DFechaNac", lstrDatosPaciente(6).Substring(6, 2) & "/" & lstrDatosPaciente(6).Substring(4, 2) & "/" & lstrDatosPaciente(6).Substring(0, 4))
            lobjImage.VirtualFields.SetFieldValue("DFechaNac", lstrDatosPaciente(6).ToString())
            lobjImage.VirtualFields.SetFieldValue("DNoHist", lstrDatosPaciente(1).ToString())
            '           lobjImage.VirtualFields.SetFieldValue("DNoSS", lstrDatosPaciente(6).ToString())
            '          lobjImage.VirtualFields.SetFieldValue("DNoHist2", lstrDatosPaciente(7).ToString())
            '         lobjImage.VirtualFields.SetFieldValue("DNoHist3", lstrDatosPaciente(8).ToString())

            'lobjImage.VirtualFields.SetFieldValue("DNoHist2", mobjOmega2.ConsultaNHUSAbyNHC(lobjImage.VirtualFields.GetFieldValue("DNoHist"), UtilGlobal.UShared.GetFolderLaboCFG & "\VirgenRocio.ini"))

            CargarDatosPaciente(lobjImage)
            ColorearVerificacion(lobjImage)

        End If

    End Sub

    Private Sub txtBuscar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscar.TextChanged

    End Sub

    Private Sub txtBuscarCustom1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscarCustom1.KeyDown

        If Me.txtBuscarCustom1.Text.Trim.Length > 0 And e.KeyCode = Keys.Enter Then
            Try
                DialogoMedico()
            Catch ex As Exception
                MessageBox.Show("Error en la consulta de médico" & vbCrLf & ex.Source & vbCrLf & ex.Message,
                                             "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        ElseIf e.KeyCode = Keys.Enter Then

            ResetMedico()

        End If

    End Sub

    ' ************************************************************************************************
    ' DialogoMedico
    ' Desc: Rutina que gestiona el diálogo de médicos
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub DialogoMedico()

        'Dim lstrResultado As String = mobjOmega2.CapturaMedicos("Búsqueda de médicos", "Texto de búsqueda", Me.txtBuscarCustom1.Text.Trim.ToUpper, True)

        'If lstrResultado.Trim.Length > 0 Then
        'Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        'CargaDatosMedicoConsulta(lstrResultado)
        '            mobjConsultasUtil.ConsultaCorrelaciones(lobjImage, mobjOmega2)
        'CargarDatosTab1(lobjImage)
        'ColorearVerificacion(lobjImage)
        'Me.txtBuscarCustom4.Focus()
        'Else
        'CampoModificado("DDoctor", "")
        'End If

    End Sub

    ' ************************************************************************************************
    ' CargaDatosMedicoConsulta
    ' Desc: Rutina que carga en los controles del verificador el resultado de la consulta de médico
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargaDatosMedicoConsulta(ByVal pstrResultado As String)

        Dim lstrDatosMedico() As String = pstrResultado.Split("|")
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        If lstrDatosMedico.Length = 2 Then
            lobjImage.VirtualFields.SetFieldValue("DDoctor", lstrDatosMedico(0).ToString())
            lobjImage.VirtualFields.SetFieldValue("DTDoctor", lstrDatosMedico(1).ToString())
            'If Not lobjImage.ARData Is Nothing Then lobjImage.VirtualFields.SetFieldValue("DDoctor", "")
            CargarDatosMedicoTab1(lobjImage)
        End If
        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub txtBuscarCustom1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscarCustom1.TextChanged

    End Sub

    Private Sub txtBuscarCustom2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscarCustom2.KeyDown


        If Me.txtBuscarCustom2.Text.Trim.Length > 0 And e.KeyCode = Keys.Enter Then
            Try
                DialogoDestino()
            Catch ex As Exception
                MessageBox.Show("Error en la consulta de destino" & vbCrLf & ex.Source & vbCrLf & ex.Message,
                                                             "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        ElseIf e.KeyCode = Keys.Enter Then
            ResetDestino()
        End If

    End Sub

    ' ************************************************************************************************
    ' DialogoDestino
    ' Desc: Rutina que gestiona el diálogo de destino
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub DialogoDestino()

        Dim lstrResultado As String = mobjOmega2.CapturaDestinos("Búsqueda de destinos", "Texto de búsqueda", Me.txtBuscarCustom2.Text.Trim.ToUpper, True)
        If lstrResultado.Trim.Length > 0 Then
            Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
            CargaDatosDestinoConsulta(lstrResultado)
            'ConsultaCorrelaciones(lobjImage, mobjOmega2, 2)
            CargarDatosTab1(lobjImage)
            ColorearVerificacion(lobjImage)
            Me.txtBuscarCustom3.Focus()
        Else
            CampoModificado("DDestino", "")
        End If

    End Sub

    ' ************************************************************************************************
    ' CargaDatosDestinoConsulta
    ' Desc: Rutina que carga en los controles del verificador el resultado de la consulta de destino
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargaDatosDestinoConsulta(ByVal pstrResultado As String)

        Dim lstrDatosDestino() As String = pstrResultado.Split("|")
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        If lstrDatosDestino.Length = 2 Then
            lobjImage.VirtualFields.SetFieldValue("DDestino", lstrDatosDestino(0).ToString())
            lobjImage.VirtualFields.SetFieldValue("TDestino", lstrDatosDestino(1).ToString())
            'If Not lobjImage.ARData Is Nothing Then lobjImage.VirtualFields.SetFieldValue("DDestino", "")
            CargarDatosDestinoTab1(lobjImage)
        End If

        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub txtBuscarCustom2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscarCustom2.TextChanged

    End Sub

    Private Sub txtBuscarCustom3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscarCustom3.KeyDown

        If Me.txtBuscarCustom3.Text.Trim.Length > 0 And e.KeyCode = Keys.Enter Then
            Try
                DialogoOrigen()
            Catch ex As Exception
                MessageBox.Show("Error en la consulta de origen" & vbCrLf & ex.Source & vbCrLf & ex.Message,
                                                             "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        ElseIf e.KeyCode = Keys.Enter Then
            ResetOrigen()
        End If

    End Sub

    ' ************************************************************************************************
    ' DialogoOrigen
    ' Desc: Rutina que gestiona el diálogo de origen
    ' NBL 16/7/2009
    ' ************************************************************************************************
    Private Sub DialogoOrigen()

        'Dim lstrResultado As String = mobjOmega2.CapturaOrigenes("Búsqueda de origenes", "Texto de búsqueda", Me.txtBuscarCustom3.Text.Trim.ToUpper, True)
        'If lstrResultado.Trim.Length > 0 Then
        '    Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        '    CargaDatosOrigenConsulta(lstrResultado)
        '    'mobjConsultasUtil.ConsultaCorrelaciones(lobjImage, mobjOmega2)
        '    CargarDatosTab1(lobjImage)
        '    ColorearVerificacion(lobjImage)
        '    Me.txtBuscarCustom2.Focus()
        'Else
        '    CampoModificado("DOrigen", "")
        'End If

    End Sub

    ' ************************************************************************************************
    ' CargaDatosOrigenConsulta
    ' Desc: Rutina que carga en los controles del verificador el resultado de la consulta de origen
    ' NBL 16/7/2009
    ' ************************************************************************************************
    Private Sub CargaDatosOrigenConsulta(ByVal pstrResultado As String)

        Dim lstrDatosOrigen() As String = pstrResultado.Split("|")
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        If lstrDatosOrigen.Length = 2 Then
            lobjImage.VirtualFields.SetFieldValue("DOrigen", lstrDatosOrigen(0).ToString())
            lobjImage.VirtualFields.SetFieldValue("TOrigen", lstrDatosOrigen(1).ToString())
            'If Not lobjImage.ARData Is Nothing Then lobjImage.ARData.ARTextFields.SetFieldValue("DOrigen", "")
            CargarDatosOrigenTab1(lobjImage)
        End If

        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub txtBuscarCustom4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtBuscarCustom4.KeyDown

        If Me.txtBuscarCustom4.Text.Trim.Length > 0 And e.KeyCode = Keys.Enter Then
            Try
                DialogoServicio()
            Catch ex As Exception
                MessageBox.Show("Error en la consulta de servicio" & vbCrLf & ex.Source & vbCrLf & ex.Message,
                                                             "Flexibar.NET", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End Try
        ElseIf e.KeyCode = Keys.Enter Then
            ResetServicio()
        End If

    End Sub

    ' ************************************************************************************************
    ' DialogoServicio
    ' Desc: Rutina que gestiona el diálogo de origen
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub DialogoServicio()

        'Dim lstrResultado As String = mobjOmega2.CapturaServicios("Búsqueda de servicios", "Texto de búsqueda", Me.txtBuscarCustom4.Text.Trim.ToUpper, True)
        'If lstrResultado.Trim.Length > 0 Then
        '    CargaDatosServicioConsulta(lstrResultado)
        '    Me.txtBuscarCustom3.Focus()
        'Else
        '    CampoModificado("DServicio", "")
        'End If

    End Sub

    ' ************************************************************************************************
    ' CargaDatosServicioConsulta
    ' Desc: Rutina que carga en los controles del verificador el resultado de la consulta de origen
    ' NBL 13/5/2009
    ' ************************************************************************************************
    Private Sub CargaDatosServicioConsulta(ByVal pstrResultado As String)

        Dim lstrDatosServicio() As String = pstrResultado.Split("|")
        Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(Me.mintIndex - 1)
        If lstrDatosServicio.Length = 2 Then
            lobjImage.VirtualFields.SetFieldValue("DServicio", lstrDatosServicio(0).ToString())
            lobjImage.VirtualFields.SetFieldValue("TServicio", lstrDatosServicio(1).ToString())
            'If Not lobjImage.ARData Is Nothing Then lobjImage.VirtualFields.SetFieldValue("DServico", "")
            CargarDatosServicioTab1(lobjImage)
        End If

        ColorearVerificacion(lobjImage)

    End Sub

    Private Sub btnResetCustom1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetCustom1.Click

        ResetMedico()

    End Sub

    Private Sub btnResetCustom2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetCustom2.Click

        ResetDestino()

    End Sub

    Private Sub btnrResetCustom3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrResetCustom3.Click

        ResetOrigen()

    End Sub

    Private Sub btnResetCustom4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetCustom4.Click

        ResetServicio()

    End Sub

    Private Sub txtObservaciones_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtObservaciones.KeyPress

        If e.KeyChar.ToString = """" Or e.KeyChar.ToString = "`" Or e.KeyChar.ToString = "´" Or e.KeyChar.ToString = "#" Or e.KeyChar.ToString = "|" Or
         e.KeyChar.ToString = "'" Or e.KeyChar.ToString = "^" Then
            e.Handled = True
        End If

    End Sub

    Private Sub txtObservaciones_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtObservaciones.Validated

        CampoModificado("DObservaciones", Me.txtObservaciones.Text.Trim)

    End Sub

    Private Sub txtDiagnosticos_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDiagnosticos.KeyPress

        If e.KeyChar.ToString = """" Or e.KeyChar.ToString = "`" Or e.KeyChar.ToString = "´" Or e.KeyChar.ToString = "#" Or e.KeyChar.ToString = "|" Or
         e.KeyChar.ToString = "'" Or e.KeyChar.ToString = "^" Then
            e.Handled = True
        End If

    End Sub

    Private Sub txtDiagnosticos_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDiagnosticos.Validated

        CampoModificado("DCDiagnostico", Me.txtDiagnosticos.Text.Trim)

    End Sub

    Private Sub txtObservaciones_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtObservaciones.TextChanged

    End Sub

    Private Sub txtSemanaGestacion_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSemanaGestacion.KeyPress

        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub txtSemanaGestacion_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSemanaGestacion.TextChanged

    End Sub

    Private Sub txtSemanaGestacion_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSemanaGestacion.Validated

        CampoModificado("D_SemanaGestacion", Me.txtSemanaGestacion.Text.Trim)

        If IsNumeric(Me.txtSemanaGestacion.Text.Trim) Then
            CambiaSeleccionMarcas("", True)
        End If

    End Sub

    Private Sub txtVolumen_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVolumen.KeyPress

        If e.KeyChar.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar.IsControl(e.KeyChar) Then
            e.Handled = False
        Else
            e.Handled = True
        End If

    End Sub

    Private Sub txtVolumen_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVolumen.TextChanged

    End Sub

    Private Sub txtVolumen_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVolumen.Validated

        CampoModificado("D_VOLUMEN_ORINA", Me.txtVolumen.Text)

    End Sub

    Private Sub txtNumeroPeticion_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNumeroPeticion.TextChanged

    End Sub

    Private Sub txtDiagnosticos_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDiagnosticos.TextChanged

    End Sub

    Private Sub txtBuscarCustom3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscarCustom3.TextChanged

    End Sub

    Private Sub txtBuscarCustom4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtBuscarCustom4.TextChanged

    End Sub

    Private Sub txtCama_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCama3.TextChanged

    End Sub

    Private Sub txtCama_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCama3.Validated

        CampoModificado("DCama", Me.txtCama3.Text.Trim)

    End Sub

    Private Sub txtNHC1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNHC1.Validated

        CampoModificado("DNoHist", Me.txtNHC1.Text.Trim)

    End Sub

    Private Sub txtNombre_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNombre.Validated

        CampoModificado("DNombre", Me.txtNombre.Text.Trim)

    End Sub

    Private Sub txtApellido1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtApellido1.Validated

        CampoModificado("DApellido1", Me.txtApellido1.Text.Trim)

    End Sub

    Private Sub txtSexo_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSexo.Validated

        CampoModificado("DSexo", Me.txtSexo.Text.Trim)

    End Sub

    Private Sub txtFechaNacimiento_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFechaNacimiento.Validated

        CampoModificado("DFechaNac", Me.txtFechaNacimiento.Text.Trim)

    End Sub

    Private Sub txtNHCCentro_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNSS.Validated

        CampoModificado("DNoHist2", Me.txtNSS.Text.Trim)

    End Sub

    Private Sub txtCodigoCustom1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCodigoCustom1.Validated

        CampoModificado("DDoctor", Me.txtCodigoCustom1.Text.Trim)

    End Sub

    Private Sub txtServicio_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtServicio.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.txtCodigoCustom2.Focus()
        End If

    End Sub

    Private Sub txtServicio_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtServicio.Validating

        CampoModificado("DDestino", Me.txtServicio.Text.Trim)
        CampoModificado("DServicio", Me.txtServicio.Text.Trim)
        Me.txtCodigoCustom2.Text = Me.txtServicio.Text.Trim

    End Sub

    Private Sub txtCama_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCama.Validating

        CampoModificado("DCama", Me.txtCama.Text.Trim)

    End Sub

    Private Sub txtCodigoCustom2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodigoCustom2.KeyDown

        If e.KeyCode = Keys.Enter Then
            Me.txtServicio.Focus()
        End If

    End Sub

    Private Sub txtCodigoCustom2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles txtCodigoCustom2.Validating

        CampoModificado("DDestino", Me.txtCodigoCustom2.Text.Trim)
        CampoModificado("DServicio", Me.txtCodigoCustom2.Text.Trim)
        Me.txtServicio.Text = Me.txtCodigoCustom2.Text.Trim

    End Sub

    Private Sub txtCodigoCustom2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodigoCustom2.TextChanged

    End Sub



    Public Sub KeysCombinationReceived(pstrModifier As String, pstrKey As String) Implements IFlexiValidator.KeysCombinationReceived
        Throw New NotImplementedException()
    End Sub
End Class

