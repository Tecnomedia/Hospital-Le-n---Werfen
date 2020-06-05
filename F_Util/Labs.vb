Imports System.Drawing
Imports System.Windows.Forms

Public Class Labs

    ' *******************************************************************************************************
    ' CreaCamposVirtualesExport
    ' Desc: Rutina que crea las variables virtuales que son necesarios para la exportación
    ' NBL 22/4/2009
    ' *******************************************************************************************************
    Public Sub CreaCamposVirtualesExport(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                        ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        Dim lobjVirtuals As LibFlexibarNETObjects.colVirtualFields = pFlexibarBatch.Images(pCodedNavigation.ImageNumber).VirtualFields

        If Not lobjVirtuals.FieldExists("DNoPet") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoPet", ""))
        If Not lobjVirtuals.FieldExists("DNoPet2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoPet2", ""))
        If Not lobjVirtuals.FieldExists("DNoPet3") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoPet3", ""))
        If Not lobjVirtuals.FieldExists("DNoHist") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoHist", ""))
        If Not lobjVirtuals.FieldExists("DNoHist2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoHist2", ""))
        If Not lobjVirtuals.FieldExists("DNoHist3") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoHist3", ""))
        If Not lobjVirtuals.FieldExists("DNoHistFusion") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoHistFusion", ""))
        If Not lobjVirtuals.FieldExists("DEpisodio") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DEpisodio", ""))
        If Not lobjVirtuals.FieldExists("DActo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DActo", ""))
        If Not lobjVirtuals.FieldExists("DNoSS") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNoSS", ""))
        If Not lobjVirtuals.FieldExists("DDNI") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDNI", ""))
        If Not lobjVirtuals.FieldExists("DApellido1") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DApellido1", ""))
        If Not lobjVirtuals.FieldExists("DApellido2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DApellido2", ""))
        If Not lobjVirtuals.FieldExists("DNombre") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNombre", ""))
        If Not lobjVirtuals.FieldExists("DFechaNac") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFechaNac", ""))
        If Not lobjVirtuals.FieldExists("DSexo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DSexo", ""))
        If Not lobjVirtuals.FieldExists("DDireccion") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDireccion", ""))
        If Not lobjVirtuals.FieldExists("DTelefono") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTelefono", ""))
        If Not lobjVirtuals.FieldExists("DPoblacion") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DPoblacion", ""))
        If Not lobjVirtuals.FieldExists("DCPostal") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DCPostal", ""))
        If Not lobjVirtuals.FieldExists("DDoctor") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDoctor", ""))
        If Not lobjVirtuals.FieldExists("DTDoctor") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTDoctor", ""))
        If Not lobjVirtuals.FieldExists("DFactura") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFactura", ""))
        If Not lobjVirtuals.FieldExists("DCDiagnostico") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DCDiagnostico", ""))
        If Not lobjVirtuals.FieldExists("DTDiagnostico") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTDiagnostico", ""))
        If Not lobjVirtuals.FieldExists("DPrioridad") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DPrioridad", ""))
        If Not lobjVirtuals.FieldExists("DCama") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DCama", ""))
        If Not lobjVirtuals.FieldExists("DTipo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTipo", ""))
        If Not lobjVirtuals.FieldExists("DMotivo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DMotivo", ""))
        If Not lobjVirtuals.FieldExists("DServicio") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DServicio", ""))
        If Not lobjVirtuals.FieldExists("DOrigen") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DOrigen", ""))
        If Not lobjVirtuals.FieldExists("DDestino") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDestino", ""))
        If Not lobjVirtuals.FieldExists("DGrupo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DGrupo", ""))
        If Not lobjVirtuals.FieldExists("DTipoFisiologico") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTipoFisiologico", ""))
        If Not lobjVirtuals.FieldExists("DFormID") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFormID", ""))
        If Not lobjVirtuals.FieldExists("DObservaciones") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DObservaciones", ""))
        If Not lobjVirtuals.FieldExists("DFHExtraccion") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFHExtraccion", ""))
        If Not lobjVirtuals.FieldExists("DFHRegistro") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFHRegistro", ""))
        If Not lobjVirtuals.FieldExists("DPruebas") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DPruebas", ""))
        If Not lobjVirtuals.FieldExists("DPerfiles") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DPerfiles", ""))
        If Not lobjVirtuals.FieldExists("DMuestra") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DMuestra", ""))
        If Not lobjVirtuals.FieldExists("DResultados") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DResultados", ""))
        If Not lobjVirtuals.FieldExists("DNTelefono") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DNTelefono", ""))
        If Not lobjVirtuals.FieldExists("DFaxResultados") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DFaxResultados", ""))
        If Not lobjVirtuals.FieldExists("DScanStation") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DScanStation", ""))
        If Not lobjVirtuals.FieldExists("DBatchNo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DBatchNo", ""))
        If Not lobjVirtuals.FieldExists("DPageNo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DPageNo", ""))
        If Not lobjVirtuals.FieldExists("DUser") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DUser", ""))

        ' NBL 21/5/2010 Añado el campo segundo doctor ya que es un campo necesario para los proyectos de Siemens
        If Not lobjVirtuals.FieldExists("DDoctor2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDoctor2", ""))
        If Not lobjVirtuals.FieldExists("DTDoctor2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DTDoctor2", ""))

        ' PVT 20/09/2010 Creamos la variable virtual DExportado que indica si una petición ha sido exportada en modo individual.

        If Not lobjVirtuals.FieldExists("DExportado") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DExportado", "0"))

        ' PVT 22/09/2010 Creamos la variable virtual DCargo

        'If Not lobjVirtuals.FieldExists("DCargo") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DCargo", "0"))



    End Sub

    ' *******************************************************************************************************
    ' CreaCamposVirtualesInternas
    ' Desc: Rutina que crea las variables virtuales que son necesarios para la gestión interna de la aplicación
    ' NBL 7/4/2010
    ' *******************************************************************************************************
    Public Sub CreaCamposVirtualesInternas(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                        ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        Dim lobjVirtuals As LibFlexibarNETObjects.colVirtualFields = pFlexibarBatch.Images(pCodedNavigation.ImageNumber).VirtualFields

        If Not lobjVirtuals.FieldExists("XAnclado") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XAnclado", "0"))
        If Not lobjVirtuals.FieldExists("XRevisado") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XRevisado", "0"))
        If Not lobjVirtuals.FieldExists("XExportado") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XExportado", "0"))
        If Not lobjVirtuals.FieldExists("XConsultaHIS") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XConsultaHIS", ""))
        If Not lobjVirtuals.FieldExists("XConsultaLIS") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XConsultaLIS", ""))

        If Not lobjVirtuals.FieldExists("DDestinoAUX") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DDestinoAUX", ""))

    End Sub

    ' ***********************************************************************************************************************************************************************
    ' CargarPruebasListView
    ' Desc: Carga en el listview que pasamos las descripciones según las pruebas que estan seleccionadas
    ' ABRV_PRUEBA_1^DESC_PRUEBA_1#ABRV_MUESTRA_1^DESC_MUESTRA_1|ABRV_PRUEBA_2^DESC_PRUEBA_2#ABRV_MUESTRA_2^DESC_MUESTRA_2
    ' NBL 22/09/2010
    ' ***********************************************************************************************************************************************************************
    Public Sub CargarPruebasListView(ByVal pstrListadoCodigos As String, ByVal pstrListadoDescripciones As String, ByRef pobjListView As ListView, ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrLista As New ArrayList
        pobjListView.Items.Clear()

        If pstrListadoCodigos.StartsWith(",") Then pstrListadoCodigos = pstrListadoCodigos.Substring(1)
        If pstrListadoCodigos.Trim.Length = 0 Then Return

        ' PVT 28/09/2010: Parche que quita los valores null ("") o los valores ("^||") que pueden llegar en la variable pstrListadoCodigos
        Dim lstrPruebasTemp() As String = pstrListadoCodigos.Split(",")
        Dim intPruebas As Integer = 0

        For i As Integer = 0 To lstrPruebasTemp.Length - 1
            If Not String.IsNullOrEmpty(lstrPruebasTemp(i)) AndAlso lstrPruebasTemp(i) <> "^||" Then
                intPruebas = intPruebas + 1
            End If
        Next

        Dim lstrPruebas(intPruebas - 1) As String
        intPruebas = 0

        For i As Integer = 0 To lstrPruebasTemp.Length - 1
            If Not String.IsNullOrEmpty(lstrPruebasTemp(i)) AndAlso lstrPruebasTemp(i) <> "^||" Then
                lstrPruebas(intPruebas) = lstrPruebasTemp(i)
                intPruebas = intPruebas + 1
            End If
        Next

        Try
            Dim lstrDescripciones() As String = pstrListadoDescripciones.Split("|")

            'If lstrPruebas.Length <= 1 Then Return

            Dim lstrCodigoPrueba As String = "", lstrTipoPrueba As String = "", lstrCodigoMuestra As String = ""
            Dim lstrAbrvPrueba As String = "", lstrDescripcionPrueba As String = "", lstrAbrvMuestra As String = "", lstrDescripcionMuestra As String = ""

            For lintContador As Integer = 0 To lstrPruebas.Length - 1
                lstrCodigoPrueba = ""
                lstrCodigoMuestra = ""
                lstrTipoPrueba = ""
                lstrAbrvPrueba = ""
                lstrDescripcionPrueba = ""
                lstrAbrvMuestra = ""
                lstrDescripcionMuestra = ""

                Dim lstrPrueba() As String = lstrPruebas(lintContador).Split("|")
                Dim lstrCodigoP() As String = lstrPrueba(0).Split("^")
                lstrCodigoPrueba = lstrCodigoP(0)

                If lstrCodigoP.Length > 1 Then
                    lstrTipoPrueba = lstrCodigoP(1)
                Else
                    lstrTipoPrueba = ""
                End If

                If lstrPrueba.Length > 1 Then
                    lstrCodigoMuestra = lstrPrueba(1)
                Else
                    lstrCodigoMuestra = ""
                End If
                If lstrCodigoPrueba.Trim.Length <> 0 Then
                    Dim lstrDescTemp() As String = lstrDescripciones(lintContador).Split("#")

                    Dim lstrDescPruebas() As String = lstrDescTemp(0).Split("^")
                    Dim lstrDescMuestras() As String = lstrDescTemp(1).Split("^")

                    lstrAbrvPrueba = lstrDescPruebas(0).Trim
                    lstrDescripcionPrueba = lstrDescPruebas(1).Trim
                    lstrAbrvMuestra = lstrDescMuestras(0).Trim
                    lstrDescripcionMuestra = lstrDescMuestras(1).Trim

                    Dim lobjListViewItem As New ListViewItem()
                    lobjListViewItem.Text = lstrCodigoPrueba

                    lobjListViewItem.SubItems.Add(lstrAbrvPrueba)
                    lobjListViewItem.SubItems.Add(lstrDescripcionPrueba)

                    If lstrCodigoMuestra <> "" Then
                        lobjListViewItem.SubItems.Add(lstrCodigoMuestra).BackColor = Color.Green
                        lobjListViewItem.SubItems.Add(lstrAbrvMuestra).BackColor = Color.Green
                        lobjListViewItem.SubItems.Add(lstrDescripcionMuestra).BackColor = Color.Green
                        lobjListViewItem.SubItems.Add(lstrTipoPrueba).BackColor = Color.Green
                    End If

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

        Catch ex As Exception

        End Try

    End Sub

    ' NUEVOP*************************************************************************************************************************************************************
    ' CalculaDescripciones
    ' Desc: Devuelve un string donde hay las descripciones encadenadas en el mismo orden que se pasan las pruebas, de esta forma
    ' ABRV_PRUEBA_1^DESC_PRUEBA_1#ABRV_MUESTRA_1^DESC_MUESTRA_1|ABRV_PRUEBA_2^DESC_PRUEBA_2#ABRV_MUESTRA_2^DESC_MUESTRA_2
    ' NBL 22/09/2010
    ' ***********************************************************************************************************************************************************************
    Public Function CalculaDescripciones(ByVal pstrListadoPruebas As String, ByVal pbolMicro As Boolean) As String

        ' Primero miramos si hay pruebas o no
        If pstrListadoPruebas.Trim.Length = 0 Then Return ""
        If pstrListadoPruebas.StartsWith(",") Then pstrListadoPruebas = pstrListadoPruebas.Substring(1)
        ' Hemos de hacer un split de comas para separar pruebas
        Dim lstrPruebas() As String = pstrListadoPruebas.Split(",")
        'If lstrPruebas.Length <= 1 Then Return ""
        Dim lstrCodigoPrueba As String = "", lstrCodigoMuestra As String = "", lstrTipoPrueba As String = ""
        Dim lstrAbrvPrueba As String = "", lstrDescripcionPrueba As String = "", lstrAbrvMuestra As String = "", lstrDescripcionMuestra As String = ""

        Dim lobjOmega As New LibOmega.clsConsulta
        Dim lstrResultado As String = ""

        lobjOmega.IniciarCargaPruebas()

        For lintContador As Integer = 0 To lstrPruebas.Length - 1

            lstrCodigoPrueba = ""
            lstrCodigoMuestra = ""
            lstrTipoPrueba = ""
            lstrAbrvPrueba = ""
            lstrDescripcionPrueba = ""
            lstrAbrvMuestra = ""
            lstrDescripcionMuestra = ""

            Dim lstrPrueba() As String = lstrPruebas(lintContador).Split("|")
            Dim lstrCodigoP() As String = lstrPrueba(0).Split("^")
            lstrCodigoPrueba = lstrCodigoP(0)
            Dim lstrBioMicro As String = ""
            If lstrCodigoP.Length > 1 Then
                lstrBioMicro = lstrCodigoP(1)
            End If

            If lstrCodigoP.Length > 1 Then
                lstrTipoPrueba = lstrCodigoP(1)
            Else
                lstrTipoPrueba = ""
            End If

            If lstrPrueba.Length > 1 Then
                lstrCodigoMuestra = lstrPrueba(1)
            Else
                lstrCodigoMuestra = ""
            End If

            If lstrCodigoPrueba.Trim.Length <> 0 Then
                If lstrBioMicro = "M" Then
                    lobjOmega.getAbrvDescripcionPruebas(lstrCodigoPrueba, lstrAbrvPrueba, lstrDescripcionPrueba, _
                                                                    lstrCodigoMuestra, lstrAbrvMuestra, lstrDescripcionMuestra, lstrTipoPrueba, True, False)
                ElseIf lstrBioMicro = "B" Then
                    lobjOmega.getAbrvDescripcionPruebas(lstrCodigoPrueba, lstrAbrvPrueba, lstrDescripcionPrueba, _
                                                                    lstrCodigoMuestra, lstrAbrvMuestra, lstrDescripcionMuestra, lstrTipoPrueba, False, False)
                Else
                    lobjOmega.getAbrvDescripcionPruebas(lstrCodigoPrueba, lstrAbrvPrueba, lstrDescripcionPrueba, _
                                                                    lstrCodigoMuestra, lstrAbrvMuestra, lstrDescripcionMuestra, lstrTipoPrueba, pbolMicro, False)
                End If
                lstrResultado &= "|" & lstrAbrvPrueba & "^" & lstrDescripcionPrueba & "#" & lstrAbrvMuestra & "^" & lstrDescripcionMuestra
            End If

        Next

        lobjOmega.CerrarCargaPruebasgaPruebas()

        If lstrResultado.Trim.Length = 0 Then Return ""

        Return lstrResultado.Substring(1)

    End Function



End Class
