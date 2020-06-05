Imports LibFlexibarNETObjects
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows.Forms

Public Class Scripts

    Implements LibFlexibarNETObjects.IFlexiScripts

    ' ****************************************************************************************************************
    ' DECLARACIÓN DE VARIABLES
    ' ****************************************************************************************************************
    Private mobjUtilLabs As New F_Util.Labs
    Private mobjUtilMarks As New F_Util.MarkToCode
    Private mobjShared As New FS_Shared.FS_Util

#Region "Entry Scripts"

    Public Sub AdvancedTransfer(ByRef pFlexibarApp As FlexibarApp,
                                ByRef pFlexibarBatch As FlexibarBatch,
                                ByRef pExportProcessResult As ExportProcessResult,
                                ByRef pFlexibarGlobal As FlexibarGlobal) Implements LibFlexibarNETObjects.IFlexiScripts.AdvancedTransfer

        TransferenciaAvanzada(pFlexibarApp, pFlexibarBatch, pExportProcessResult)

    End Sub

    Public Function ApplyAdvancedRecognition(ByVal pSeparatorBarcodes As colCoreBarcodes,
                                             ByVal pDataBarcodes As colCoreBarcodes,
                                             ByRef pFlexibarApp As FlexibarApp,
                                             ByRef pFlexibarBatch As FlexibarBatch,
                                             ByVal pCodedNavigation As CodedNavigation,
                                             ByRef pFlexibarGlobal As FlexibarGlobal) As Boolean Implements LibFlexibarNETObjects.IFlexiScripts.ApplyAdvancedRecognition

    End Function

    Public Function DocumentFileName(ByRef pImageToExport As ImageToExport,
                                     ByRef pExportProcessResult As ExportProcessResult,
                                     ByRef pFlexibarApp As FlexibarApp,
                                     ByRef pFlexibarBatch As FlexibarBatch,
                                     ByRef pFlexibarGlobal As FlexibarGlobal) As String Implements LibFlexibarNETObjects.IFlexiScripts.DocumentFileName

    End Function

    Public Function DocumentSubDirectory(ByRef pImageToExport As ImageToExport,
                                         ByRef pExportProcessResult As ExportProcessResult,
                                         ByRef pFlexibarApp As FlexibarApp,
                                         ByRef pFlexibarBatch As FlexibarBatch,
                                         ByRef pFlexibarGlobal As FlexibarGlobal) As String Implements IFlexiScripts.DocumentSubDirectory
        Throw New NotImplementedException()
    End Function

    Public Function ImageFileName(ByRef pImageToExport As ImageToExport,
                                  ByRef pExportProcessResult As ExportProcessResult,
                                  ByRef pFlexibarApp As FlexibarApp,
                                  ByRef pFlexibarBatch As FlexibarBatch,
                                  ByRef pFlexibarGlobal As FlexibarGlobal) As String Implements LibFlexibarNETObjects.IFlexiScripts.ImageFileName

    End Function

    Public Function ImageSubDirectory(ByRef pImageToExport As ImageToExport,
                                      ByRef pExportProcessResult As ExportProcessResult,
                                      ByRef pFlexibarApp As FlexibarApp,
                                      ByRef pFlexibarBatch As FlexibarBatch,
                                      ByRef pFlexibarGlobal As FlexibarGlobal) As String Implements IFlexiScripts.ImageSubDirectory
        Throw New NotImplementedException()
    End Function

    Public Sub ImportScript(ByRef pFlexibarApp As FlexibarApp,
                            ByRef pFlexibarBatch As FlexibarBatch,
                            ByRef pFileList As ArrayList,
                            ByRef pbolCaraAB As Boolean,
                            ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.ImportScript
        Throw New NotImplementedException()
    End Sub

    Public Sub NextImage(ByRef pFlexibarApp As FlexibarApp,
                         ByRef pFlexibarBatch As FlexibarBatch,
                         ByRef pCodedNavigation As CodedNavigation,
                         ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.NextImage
        Throw New NotImplementedException()
    End Sub

    Public Sub PostEvaluationScript(ByRef pFlexibarBatch As FlexibarBatch,
                                    ByRef pCodedNavigation As CodedNavigation,
                                    ByRef pFlexibarApp As FlexibarApp,
                                    ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.PostEvaluationScript

        RutinaPrincipalPostEvaluationScript(pFlexibarBatch, pCodedNavigation, pFlexibarApp)

    End Sub

    Public Sub PostFileCopyEvent(ByRef pImageToExport As ImageToExport,
                                 ByRef pExportPath As String,
                                 ByRef pExportProcessResult As ExportProcessResult,
                                 ByRef pFlexibarApp As FlexibarApp,
                                 ByRef pFlexibarBatch As FlexibarBatch,
                                 ByRef pFlexibarGlobal As FlexibarGlobal) Implements LibFlexibarNETObjects.IFlexiScripts.PostFileCopyEvent

    End Sub

    Public Sub PreviousImage(ByRef pFlexibarApp As FlexibarApp,
                             ByRef pFlexibarBatch As FlexibarBatch,
                             ByRef pCodedNavigation As CodedNavigation,
                             ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.PreviousImage
        Throw New NotImplementedException()
    End Sub

    Public Sub ScanControlEvent(ByRef pFlexibarApp As FlexibarApp,
                                ByRef pFlexibarBatch As FlexibarBatch,
                                ByRef pFlexibarGlobal As FlexibarGlobal) Implements LibFlexibarNETObjects.IFlexiScripts.ScanControlEvent

    End Sub

    Public Function TransferPrevalidation(ByRef pFlexibarApp As FlexibarApp,
                                          ByRef pFlexibarBatch As FlexibarBatch,
                                          ByRef pCodedNavigation As CodedNavigation,
                                          ByRef pFlexibarGlobal As FlexibarGlobal) As Boolean Implements IFlexiScripts.TransferPrevalidation
        Return ValidarTransferencia(pFlexibarApp, pFlexibarBatch, pCodedNavigation)
    End Function

    Public Sub AsyncTransfer(ByRef pFlexibarApp As FlexibarApp,
                             ByRef pFlexibarBatch As FlexibarBatch,
                             ByRef pExportProcessResult As ExportProcessResult,
                             ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.AsyncTransfer
        Throw New NotImplementedException()
    End Sub

    Public Sub EndAppScript(ByRef pFlexibarApp As FlexibarApp,
                            ByRef pFlexibarBatch As FlexibarBatch,
                            ByRef pobjArguments As EndAppScriptArguments,
                            ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.EndAppScript
        Throw New NotImplementedException()
    End Sub

    Public Function InitAppScript(ByRef pFlexibarApp As FlexibarApp,
                                  ByRef pFlexibarBatch As FlexibarBatch,
                                  ByRef pFlexibarGlobal As FlexibarGlobal) As Boolean Implements IFlexiScripts.InitAppScript
        Throw New NotImplementedException()
    End Function

    Public Sub KeyEventScript(ByRef pFlexibarApp As FlexibarApp,
                              ByRef pFlexibarBatch As FlexibarBatch,
                              ByRef pobjArguments As KeyEventScriptArguments,
                              ByRef pFlexibarGlobal As FlexibarGlobal) Implements IFlexiScripts.KeyEventScript
        Throw New NotImplementedException()
    End Sub

#End Region

    ' *******************************************************************************************************
    ' ValidarTransferencia
    ' Desc: Hacemos un bucle 
    ' NBL 19/5/2009
    ' *******************************************************************************************************
    Private Function ValidarTransferencia(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation) As Boolean

        Dim ldtFecha As Date = pFlexibarBatch.mobjBatchValues.BatchDate

        'If MessageBox.Show(String.Format("Los siguientes volantes se registrarán en la fecha {0}", ldtFecha.ToShortDateString) & vbCrLf & "¿Desea cambiar la fecha de registro?", "Transferencia", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = DialogResult.Yes Then
        '    Return False
        'End If

        ' Hacemos un bucle por todas las imágenes que haya en el batch, comprobamos que no tengan número de petición y que no estén marcadas para eliminar
        Dim lintNumeroImagenesBatch As Integer = pFlexibarBatch.Images.Count

        For lintContador As Integer = 0 To lintNumeroImagenesBatch - 1
            Dim lobjImage As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(lintContador)
            If Not Regex.IsMatch(lobjImage.VirtualFields.GetFieldValue("DNoPet"), "^\d{8}$") And lobjImage.RemovePage = False Then
                MsgBox("La imagen " & (lintContador + 1).ToString() & " no tiene número de petición válido", MsgBoxStyle.Exclamation, "Flexibar.NET")
                pCodedNavigation.ImageNumber = lintContador + 1
                Return False
            End If
            ' Si llega aquí es válido, lo que hago es crear las variables DFAguda y DFConvaleciente en el caso de que sea un 68
            'CamposPropios68(lobjImage)
        Next

        Return True

    End Function

    ' *******************************************************************************************************
    ' RutinaPrincipalPostEvaluationScript
    ' Desc: Rutina principal del Evaluation Script (se ejecuta por cada imagen digitalizada)
    ' NBL 21/9/2011
    ' *******************************************************************************************************
    Private Sub RutinaPrincipalPostEvaluationScript(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                     ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                     ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        UtilGlobal.UShared.EscribeLog("Inicio de Evaluation Script")

        Dim lobjImage As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(pCodedNavigation.ImageNumber)
        Dim lbolValidaMarcas As Boolean = mobjUtilMarks.ControlCalidadPlantilla(lobjImage.ARData.ARCheckmarkFields, lobjImage.ARData.TemplateName)

        UtilGlobal.UShared.EscribeLog("Control de calidad: " & lbolValidaMarcas.ToString())
        UtilGlobal.UShared.EscribeLog("Creación de variables virtuales")

        ' Creamos las variables virtuales
        mobjUtilLabs.CreaCamposVirtualesExport(pFlexibarBatch, pCodedNavigation, pFlexibarApp)
        mobjUtilLabs.CreaCamposVirtualesInternas(pFlexibarBatch, pCodedNavigation, pFlexibarApp)
        ' Creamos las variables virtuales auxiliares
        CreaCamposVirtualesAux(pFlexibarBatch, pCodedNavigation, pFlexibarApp)

        ' Leemos y acondicionamos las variables
        UtilGlobal.UShared.EscribeLog("Acondicionamos variables")
        AcondicionaVariables(pFlexibarBatch, pCodedNavigation, pFlexibarApp, lbolValidaMarcas)

        UtilGlobal.UShared.EscribeLog("Consultamos datos")

        lobjImage.VirtualFields.SetFieldValue("DPrioridad", "ER")

        Try
            ConsultaDatosFormulario(pFlexibarBatch, pCodedNavigation, pFlexibarApp)
        Catch ex As Exception
            MsgBox(ex.Source & vbCrLf & ex.Message, MsgBoxStyle.Exclamation, "Flexibar.NET")
        End Try

        CreaCampoVirtualCodigosPrueba(pFlexibarBatch, pCodedNavigation, pFlexibarApp, lbolValidaMarcas)

        If lbolValidaMarcas Then
            lobjImage.VirtualFields.SetFieldValue("XAnclado", "1")
        Else
            lobjImage.VirtualFields.SetFieldValue("XAnclado", "0")
        End If

        '' Script que tengo para escribir los txt de pruebas y/o muestras -----------------------------------------------------------------------------------------------------------
        'If lobjImage.ARData.TemplateName Is Nothing Then Exit Sub
        'If lobjImage.ARData.TemplateName = "" Then Exit Sub

        'Dim lobjStreamWriter As New IO.StreamWriter("C:\Documents and Settings\Nicolas\Mis documentos\Desarrollo\Izasa\Proyecto León Urgencias\txtMarcas\" & _
        '                                                                            lobjImage.ARData.TemplateName & ".txt")
        'lobjStreamWriter.WriteLine("#")
        '' MessageBox.Show("Hola")
        'For Each lobjCheckMark As LibFlexibarNETObjects.ARCheckmarkField In lobjImage.ARData.ARCheckmarkFields
        '    If Regex.IsMatch(lobjCheckMark.Name, "^[M,P]\d{4}[M,P,L]") Then
        '        'If Regex.IsMatch(lobjCheckMark.Name, "^M19") Then
        '        lobjStreamWriter.WriteLine(lobjCheckMark.Name.Substring(0, 6) & "#" & lobjCheckMark.Name & "#" & ",^B||")
        '        'lobjStreamWriter.WriteLine(lobjCheckMark.Name & "#")
        '    End If
        'Next
        'lobjStreamWriter.Close()
        '' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------

    End Sub

    ' *******************************************************************************************************
    ' CreaCampoVirtualCodigosPrueba
    ' Desc: Rutina que crea los campos virtuales que almacenan los códigos de prueba
    ' NBL 4/5/2009
    ' *******************************************************************************************************
    Private Sub CreaCampoVirtualCodigosPrueba(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                        ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                        ByVal pbolValidaMarcas As Boolean)

        Dim lobjImage As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(pCodedNavigation.ImageNumber)

        If Not lobjImage.VirtualFields.FieldExists("ACodigosMarcas") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosMarcas", ""))
        If Not lobjImage.VirtualFields.FieldExists("ACodigosManuscrito") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosManuscrito", ""))

        If Not lobjImage.VirtualFields.FieldExists("ACodigosMarcasDesc") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosMarcasDesc", ""))
        If Not lobjImage.VirtualFields.FieldExists("ACodigosManuscritoDesc") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosManuscritoDesc", ""))

        If Not lobjImage.VirtualFields.FieldExists("ACodigosMarcasHemato") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosMarcasHemato", ""))
        If Not lobjImage.VirtualFields.FieldExists("ACodigosMarcasHematoDesc") Then lobjImage.VirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosMarcasHematoDesc", ""))

        If lobjImage.ARData Is Nothing Then Return
        If lobjImage.ARData.TemplateName Is Nothing Then Return
        If lobjImage.ARData.TemplateName = "" Then Return
        If Not pbolValidaMarcas Then Return

        If lobjImage.ARData.TemplateName = "LE_URGENCIAS" Then

            ' Calculamos las pruebas de hematología
            mobjUtilMarks.CalculaPruebasHemato(lobjImage)

            ' Calculamos las pruebas de bioquímica y micro de omega
            lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", mobjUtilMarks.MarkTraductor(lobjImage.ARData.ARCheckmarkFields, lobjImage.ARData.TemplateName, False, 1, ""))
            lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), True))

            'mobjUtilMarks.CalculaPruebasBioquimica(lobjImage)

        End If

        'If lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA" Or lobjImage.ARData.TemplateName = "VP_BACTERIOLOGIA_JARA_2" Then
        '    ' En este caso la construcción de las pruebas se hace mediante la BBDD
        '    Dim lstrPruebas As String = ""
        '    Dim lstrObs As String = ""
        '    mobjShared.CalculaPruebasMicroVirgenPuerto(lobjImage, lstrPruebas, lstrObs)

        '    ' NBL 17/11/2010 Si es urocultivo hay que añadir una prueba de bioquímica
        '    If lstrPruebas.Contains(",22^M|76|") Then
        '        lstrPruebas &= ",9055^B||"
        '    End If

        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", lstrPruebas)
        '    lobjImage.VirtualFields.SetFieldValue("DObservaciones2", lstrObs)
        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), True))
        'Else
        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", mobjUtilMarks.MarkTraductor(lobjImage.ARData.ARCheckmarkFields, lobjImage.ARData.TemplateName, 1))
        '    lobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", mobjUtilLabs.CalculaDescripciones(lobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"), False))
        'End If

    End Sub



    ' *******************************************************************************************************
    ' ConsultaDatosFormulario
    ' Desc: Rutina para consultar 
    ' NBL 28/4/2009
    ' *******************************************************************************************************
    Private Sub ConsultaDatosFormulario(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                        ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        Dim lobjImage As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(pCodedNavigation.ImageNumber)

        ' Hacemos todas las consultas que sean necesarias siempre que haya identificado con una plantilla
        'If lobjImage.ARData Is Nothing Then Return
        'If lobjImage.ARData.TemplateName Is Nothing Then Return
        'If lobjImage.ARData.TemplateName = "" Then Return

        '        Dim lobjBBDD As New F_BBDD.Consultas
        Dim lobjOmega As New LibOmega.ComClsOmega

        ' PATIENT
        UtilGlobal.UShared.EscribeLog("Consultamos pacientes")
        ConsultaPacienteHIS(lobjImage)

        'TratamientoDatosDemograficos(lobjImage, lobjOmega, pFlexibarBatch)

        ' DOCTOR, SERVICIO Y ORIGEN
        UtilGlobal.UShared.EscribeLog("Consultamos doctor")
        Try
            '   TratamientoMedicoServicioOrigen(lobjImage, lobjOmega)
        Catch ex As Exception

        End Try
        ' SERVICIO
        '      ConsultaServicio(lobjImage, lobjBBDD)
        ' DESTINO
        UtilGlobal.UShared.EscribeLog("Consultamos destino")
        mobjShared.ConsultaDestino(lobjImage, lobjOmega)
        ' CORRELACIONES
        UtilGlobal.UShared.EscribeLog("Consultamos correlaciones")
        'mobjShared.ConsultaCorrelaciones(lobjImage, lobjOmega)

        '   lobjBBDD.Dispose()

    End Sub

    ' *********************************************************************************************************
    ' ConsultaPacienteHIS
    ' Desc.: Hacemos la consulta al HIS del hospital de los datos demográficos del paciente a partir del número de historia clínica
    ' NBL 6/10/2011
    ' *********************************************************************************************************
    Private Sub ConsultaPacienteHIS(ByRef pobjImage As LibFlexibarNETObjects.Image)

        If pobjImage.VirtualFields.GetFieldValue("DNoHist").Trim.Length > 0 Then
            mobjShared.RecuperaDatosPaciente(pobjImage, pobjImage.VirtualFields.GetFieldValue("DNoHist"))
        End If

    End Sub

    ' *******************************************************************************************************
    ' AcondicionaVariables
    ' Desc.: Rutina que crea las variables virtuales auxiliares
    ' NBL 22/09/2011
    ' *******************************************************************************************************
    Private Sub AcondicionaVariables(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                     ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                     ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                     ByVal pbolValidaMarcasValidacion As Boolean)

        AcondicionaCodigosBarra(pFlexibarBatch, pCodedNavigation, pFlexibarApp)

        Dim lobjImage As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(pCodedNavigation.ImageNumber)
        mobjUtilMarks.CamposNivelNegro(lobjImage)

        If Not pbolValidaMarcasValidacion Then Exit Sub

        lobjImage.VirtualFields.SetFieldValue("DCDiagnostico", lobjImage.ARData.ARTextFields.GetFieldValue("AUX_DIAGNOSTICO"))
        lobjImage.VirtualFields.SetFieldValue("DDoctor", mobjUtilMarks.ArrayTraductor(lobjImage.ARData.ARCheckmarkFields, "M03", False, lobjImage.ARData.TemplateName))
        lobjImage.VirtualFields.SetFieldValue("DDestinoAUX", UtilGlobal.UShared.QuitarCerosIzquierda(mobjUtilMarks.ArrayTraductor(lobjImage.ARData.ARCheckmarkFields, "M01", False, lobjImage.ARData.TemplateName)))
        lobjImage.VirtualFields.SetFieldValue("DServicio", UtilGlobal.UShared.QuitarCerosIzquierda(mobjUtilMarks.ArrayTraductor(lobjImage.ARData.ARCheckmarkFields, "M01", False, lobjImage.ARData.TemplateName)))

    End Sub

    ' *******************************************************************************************************
    ' AcondicionaCodigosBarra
    ' Desc.: Acondicionamos los códigos de barra, con los datos que haya
    ' NBL 22/09/2011
    ' ********************************************************************************************************
    Private Sub AcondicionaCodigosBarra(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                            ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                            ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        Dim lobjVirtuals As LibFlexibarNETObjects.colVirtualFields = pFlexibarBatch.Images(pCodedNavigation.ImageNumber).VirtualFields

        ' El número de petición está definido como el código de barras de contenido
        If pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData.Count > 0 Then

            For lintContador As Integer = 0 To pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData.Count - 1

                If pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).BarcodeType = Tasman.Bars.Symbologies.Codabar And _
                    pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.Length = 10 Then

                    lobjVirtuals.SetFieldValue("DNoPet", pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.Substring(1, pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.Length - 2))

                End If

                ' NBL 14/11/2012 Modificamos la lectura de códigos de barra de historia clínica para que los códigos de barra vayan de longitud de 3 a 7.
                If pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).BarcodeType = Tasman.Bars.Symbologies.Code128 And _
                    (pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.Length >= 3 And _
                     pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.Length <= 7) Then

                    lobjVirtuals.SetFieldValue("DNoHist", pFlexibarBatch.Images(pCodedNavigation.ImageNumber).CoreData(lintContador).Value.ToString())

                End If

            Next

        End If

    End Sub

    ' *******************************************************************************************************
    ' CreaCamposVirtualesAux
    ' Desc: Rutina que crea las variables virtuales auxiliares
    ' NBL 27/8/2010
    ' *******************************************************************************************************
    Private Sub CreaCamposVirtualesAux(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pCodedNavigation As LibFlexibarNETObjects.CodedNavigation, _
                                                        ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp)

        ' Creo los auxiliares que tendrán la descripcion de los que me interesan
        Dim lobjVirtuals As LibFlexibarNETObjects.colVirtualFields = pFlexibarBatch.Images(pCodedNavigation.ImageNumber).VirtualFields

        If Not lobjVirtuals.FieldExists("TServicio") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("TServicio", ""))
        If Not lobjVirtuals.FieldExists("TOrigen") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("TOrigen", ""))
        If Not lobjVirtuals.FieldExists("TDestino") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("TDestino", ""))

        If Not lobjVirtuals.FieldExists("XDemograficos") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("XDemograficos", "0"))
        ' Variable virtual para almacenar las observaciones de pruebas de micro, no se muestra pero al hacer la transferencia
        ' se tienen que unir con las que se muestran en el verificador
        If Not lobjVirtuals.FieldExists("DObservaciones2") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("DObservaciones2", ""))

        ' Variable donde pondremos el churro del código de barras
        If Not lobjVirtuals.FieldExists("A_PDF417") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("A_PDF417", ""))

        ' Creo los campos virtuales para la orina y la semana de gestación
        If Not lobjVirtuals.FieldExists("D_VOLUMEN_ORINA") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("D_VOLUMEN_ORINA", ""))
        If Not lobjVirtuals.FieldExists("D_SemanaGestacion") Then lobjVirtuals.Add(New LibFlexibarNETObjects.VirtualField("D_SemanaGestacion", ""))

    End Sub

    ' **************************************************************************************************************
    ' TransferenciaAvanzada
    ' Desc.: 
    ' NBL 27/9/2011
    ' **************************************************************************************************************
    Private Sub TransferenciaAvanzada(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                            ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                            ByRef pExportProcessResult As LibFlexibarNETObjects.ExportProcessResult)

        ' Hay dos transferencias que realizaremos en paralelo
        Dim lobjModulabASTM As New ASTM_MODULAB.Export
        Dim lobjOmegaASTM As New ASTM_OMEGA.Constructor

        Dim lobjINI As New UtilGlobal.clsINI

        Dim lstrExportMODULAB As String = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\LEON_URGENCIAS.ini", "General", "Export_MODULAB", "0")
        Dim lstrExportOMEGA As String = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\LEON_URGENCIAS.ini", "General", "Export_OMEGA", "0")

        If lstrExportMODULAB = "1" Then
            Try
                lobjModulabASTM.ExportASTM(pFlexibarApp, pFlexibarBatch, pExportProcessResult)
            Catch ex As Exception
                pExportProcessResult.ExportStatus = enExportStatus.ExportError
                pExportProcessResult.ErrorDescription = ex.Message
                Exit Sub
            End Try
        End If

        If lstrExportOMEGA = "1" Then
            Try
                lobjOmegaASTM.ExportarASTM(pFlexibarApp, pFlexibarBatch, pExportProcessResult)
            Catch ex As Exception
                pExportProcessResult.ExportStatus = enExportStatus.ExportError
                pExportProcessResult.ErrorDescription = ex.Message
            End Try
        End If

    End Sub








End Class
