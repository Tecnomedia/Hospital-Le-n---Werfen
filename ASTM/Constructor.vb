Imports System.IO
Imports System.Text.RegularExpressions

Public Class Constructor

    ' Ruta archivo ini
    Private mstrRutaINI As String = UtilGlobal.UShared.GetFolderLaboCFG & "\OMEGA_ASTM.ini"
    Private mobjINI As New UtilGlobal.clsINI

    ' Objeto LibImage
    Private mobjImage As New LibImage.Utils

    Private mstrCadenaFinal As String = ""
    ' OJO, tendré que tener en cuenta el tipo de prueba por defecto
    Private mstrTipoPruebaPorDefecto As String = ""

    Private mbolExisteConfiguracionParticular As Boolean = False

    ' Variables de la configuración de la exportación ----------------------------------------------------------------------
    Private mbolActivarExportacionASTM As Boolean = False
    Private mstrExtensionASTM As String = ""
    Private mbolActivarDebug As Boolean = False
    Private mstrRutaBaseDebug As String = ""
    Private mbolInteractiva As Boolean = False
    ' NBL 20/11/2009
    Private mbolMMDD As Boolean
    ' ------------------------------------------------------------------------------------------------------------------------
    ' Configuración standard
    Private mbolStdActivada As Boolean = False
    Private mstrStdRutaBaseExportacion As String = ""
    Private mstrStdRutaBaseImagen As String = ""
    Private mbolStdActivarExportacionImagen As Boolean = False
    Private mbolStdSubdirectorioFecha As Boolean = False
    Private mintStdFecha As Integer = 0
    Private mbolStdAnexarImagenes As Boolean = False

    ' 20/09/2010: Variables de módulo añadidas para la parametrización del formato de la fecha de Registro/Extracción en el ASTM para la configuración estandard.
    Private mbolStdBatchDateToFRegistro As Boolean = False
    Private mbolStdBatchDateToFExtraccion As Boolean
    Private mbolStdAddHoraToFRegistro As Boolean = False
    Private mbolStdAddHoraToFExtraccion As Boolean = False
    ' ------------------------------------------------------------------------------------------------------------------------
    ' Configuración standard backup
    Private mbolStdActivadaBackup As Boolean = False
    Private mstrStdRutaBaseExportacionBackup As String = ""
    Private mbolStdActivarExportacionImagenBackup As Boolean = False
    Private mstrStdRutaBaseImagenBackup As String = ""
    Private mbolStdSubdirectorioFechaBackup As Boolean = False
    Private mintStdFechaBackup As Integer = 0
    Private mbolStdAnexarImagenesBackup As Boolean = False
    ' ------------------------------------------------------------------------------------------------------------------------
    ' Configuración por plantilla
    Private mbolPrtActivada As Boolean = False
    Private mstrPrtRutaBaseExportacion As String = ""
    Private mbolPrtActivarExportacionImagen As Boolean = False
    Private mstrPrtRutaBaseImagen As String = ""
    Private mbolPrtSubdirectorioFecha As Boolean = False
    Private mintPrtFecha As Integer = 0
    Private mbolPrtAnexarImagenes As Boolean = False

    ' 20/09/2010: Variables de módulo añadidas para la parametrización del formato de la fecha de Registro/Extracción en el ASTM para la configuración personalizada de plantilla.
    Private mbolPrtBatchDateToFRegistro As Boolean = False
    Private mbolPrtBatchDateToFExtraccion As Boolean
    Private mbolPrtAddHoraToFRegistro As Boolean = False
    Private mbolPrtAddHoraToFExtraccion As Boolean = False
    ' ------------------------------------------------------------------------------------------------------------------------
    ' Configuración por plantilla backup
    Private mbolPrtActivadaBackup As Boolean = False
    Private mstrPrtRutaBaseExportacionBackup As String = ""
    Private mbolPrtActivarExportacionImagenBackup As Boolean = False
    Private mstrPrtRutaBaseImagenBackup As String = ""
    Private mbolPrtSubdirectorioFechaBackup As Boolean = False
    Private mintPrtFechaBackup As Integer = 0
    Private mbolPrtAnexarImagenesBackup As Boolean = False

    Private mstrNoPet As String

    Private mbolNombreImagenMMDD As Boolean = False

    ' **********************************************************************************************
    ' ExportarASTM
    ' Desc.: Rutina principal de exportación de la cadena ASTM
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Public Overloads Sub ExportarASTM(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                            ByVal pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                            ByRef pExportProcessResult As LibFlexibarNETObjects.ExportProcessResult)

        ' En primer hemos de mirar si existe el archivo de configuración
        If Not My.Computer.FileSystem.FileExists(mstrRutaINI) Then
            MsgBox("El archivo de configuración no existe." & vbCrLf & "Consulte su administrador", MsgBoxStyle.Exclamation, "ASTM export")
            pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
            pExportProcessResult.ErrorDescription = "El archivo de configuración no existe. Consulte su administrador"
            Exit Sub
        End If

        Try

            InicializarExportASTM()

            ' Si no está activada la exportación salimos y no hacemos nada
            If Not Me.mbolActivarExportacionASTM Then Exit Sub


            ' Bucle por cada una de las peticiones
            For lintContador As Integer = 0 To pFlexibarBatch.Images.Count - 1

                Dim lobjImagen As LibFlexibarNETObjects.Image = pFlexibarBatch.Images(lintContador)

                If Not lobjImagen.RemovePage And (lobjImagen.VirtualFields.GetFieldValue("ACodigosMarcas").Trim.Length > 0 Or lobjImagen.VirtualFields.GetFieldValue("ACodigosManuscrito").Trim.Length > 0) Then

                    CompruebaExportacionParticularPlantilla(lobjImagen)

                    ' Por cada imagen nueva que vaya a tratar, tengo que resetear la cadena final
                    Me.mstrCadenaFinal = ""

                    CreateStringASTM(lobjImagen, pFlexibarBatch.mobjBatchValues.BatchDate)
                    If Me.mbolInteractiva Then MsgBox(Me.mstrCadenaFinal)

                    If Me.mbolExisteConfiguracionParticular Then
                        If Me.mbolPrtActivada Then ExportarArchivoASTM(lobjImagen, Me.mstrPrtRutaBaseExportacion)
                        If Me.mbolPrtActivarExportacionImagen Then ExportarImagen(lobjImagen, Me.mstrPrtRutaBaseImagen, Me.mbolPrtSubdirectorioFecha,
                                                                                        Me.mintPrtFecha, Me.mbolPrtAnexarImagenes, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                        If Me.mbolPrtActivadaBackup Then ExportarArchivoASTM(lobjImagen, Me.mstrPrtRutaBaseExportacionBackup)
                        If Me.mbolPrtActivarExportacionImagenBackup Then ExportarImagen(lobjImagen, Me.mstrPrtRutaBaseImagenBackup, Me.mbolPrtSubdirectorioFechaBackup,
                                                                                        Me.mintPrtFechaBackup, Me.mbolPrtAnexarImagenesBackup, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                    Else
                        If Me.mbolStdActivada Then ExportarArchivoASTM(lobjImagen, Me.mstrStdRutaBaseExportacion)
                        If Me.mbolStdActivarExportacionImagen Then ExportarImagen(lobjImagen, Me.mstrStdRutaBaseImagen, Me.mbolStdSubdirectorioFecha,
                                                                                        Me.mintStdFecha, Me.mbolStdAnexarImagenes, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                        If Me.mbolStdActivadaBackup Then ExportarArchivoASTM(lobjImagen, Me.mstrStdRutaBaseExportacionBackup)
                        If Me.mbolStdActivarExportacionImagenBackup Then ExportarImagen(lobjImagen, Me.mstrStdRutaBaseImagenBackup, Me.mbolStdSubdirectorioFechaBackup,
                                                                                        Me.mintStdFechaBackup, Me.mbolStdAnexarImagenesBackup, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                    End If

                    If Me.mbolActivarDebug And My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug) Then
                        If Not My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug & "\img") Then My.Computer.FileSystem.CreateDirectory(Me.mstrRutaBaseDebug & "\img")
                        If Not My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug & "\pet") Then My.Computer.FileSystem.CreateDirectory(Me.mstrRutaBaseDebug & "\pet")
                        ExportarArchivoASTM(lobjImagen, Me.mstrRutaBaseDebug & "\pet", pFlexibarBatch.mobjBatchValues.BatchDate)
                        ExportarImagen(lobjImagen, Me.mstrRutaBaseDebug & "\img", True, 2, True, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                    End If

                    LogPostEvaluationScript(lobjImagen)

                End If

            Next

        Catch ex As Exception

            pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
            pExportProcessResult.ErrorDescription = ex.Message.ToString()

        End Try

    End Sub


    ' **********************************************************************************************
    ' ExportarASTM (Método sobrecargado)
    ' Desc.: Rutina principal sobrecargada de exportación de la cadena ASTM para la transferencia individual de una imagen. Se tiene que pasar como parámetro el índice de la imagen seleccionada.
    ' NBL 6/10/2009
    ' **********************************************************************************************

    Public Overloads Function ExportarASTM(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pintIndex As Integer, ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                            ByVal pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                            ByRef pstrExportMessage As String) As Boolean

        ' En primer hemos de mirar si existe el archivo de configuración
        If Not My.Computer.FileSystem.FileExists(mstrRutaINI) Then
            MsgBox("El archivo de configuración no existe." & vbCrLf & "Consulte su administrador", MsgBoxStyle.Exclamation, "ASTM export")
            pstrExportMessage = "El archivo de configuración no existe. Consulte su administrador"
            Return False
        End If

        Try

            InicializarExportASTM()

            ' Si no está activada la exportación salimos y no hacemos nada
            If Not Me.mbolActivarExportacionASTM Then
                pstrExportMessage = "La transferencia de peticiones no está activada"
                Return False
            End If

            Dim lobjImagen As LibFlexibarNETObjects.Image = pobjImage

            If Not lobjImagen.RemovePage Then

                CompruebaExportacionParticularPlantilla(lobjImagen)

                ' Por cada imagen nueva que vaya a tratar, tengo que resetear la cadena final
                Me.mstrCadenaFinal = ""

                CreateStringASTM(lobjImagen, pFlexibarBatch.mobjBatchValues.BatchDate)
                If Me.mbolInteractiva Then MsgBox(Me.mstrCadenaFinal)

                If Me.mbolExisteConfiguracionParticular Then
                    ExportarArchivoASTM(lobjImagen, Me.mstrPrtRutaBaseExportacion)
                    If Me.mbolPrtActivarExportacionImagen Then ExportarImagen(lobjImagen, Me.mstrPrtRutaBaseImagen, Me.mbolPrtSubdirectorioFecha,
                                                                                    Me.mintPrtFecha, Me.mbolPrtAnexarImagenes, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                    If Me.mbolPrtActivadaBackup Then ExportarArchivoASTM(lobjImagen, Me.mstrPrtRutaBaseExportacionBackup, pFlexibarBatch.ParentFolder)
                    If Me.mbolPrtActivarExportacionImagenBackup Then ExportarImagen(lobjImagen, Me.mstrPrtRutaBaseImagenBackup, Me.mbolPrtSubdirectorioFechaBackup,
                                                                                    Me.mintPrtFechaBackup, Me.mbolPrtAnexarImagenesBackup, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                Else
                    ExportarArchivoASTM(lobjImagen, Me.mstrStdRutaBaseExportacion)
                    If Me.mbolStdActivarExportacionImagen Then ExportarImagen(lobjImagen, Me.mstrStdRutaBaseImagen, Me.mbolStdSubdirectorioFecha,
                                                                                    Me.mintStdFecha, Me.mbolStdAnexarImagenes, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                    If Me.mbolStdActivadaBackup Then ExportarArchivoASTM(lobjImagen, Me.mstrStdRutaBaseExportacionBackup)
                    If Me.mbolStdActivarExportacionImagenBackup Then ExportarImagen(lobjImagen, Me.mstrStdRutaBaseImagenBackup, Me.mbolStdSubdirectorioFechaBackup,
                                                                                    Me.mintStdFechaBackup, Me.mbolStdAnexarImagenesBackup, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                End If

                If Me.mbolActivarDebug And My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug) Then
                    If Not My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug & "\img") Then My.Computer.FileSystem.CreateDirectory(Me.mstrRutaBaseDebug & "\img")
                    If Not My.Computer.FileSystem.DirectoryExists(Me.mstrRutaBaseDebug & "\pet") Then My.Computer.FileSystem.CreateDirectory(Me.mstrRutaBaseDebug & "\pet")
                    ExportarArchivoASTM(lobjImagen, Me.mstrRutaBaseDebug & "\pet", pFlexibarBatch.mobjBatchValues.BatchDate)
                    ExportarImagen(lobjImagen, Me.mstrRutaBaseDebug & "\img", True, 2, True, pFlexibarBatch.mobjBatchValues.BatchDate, pFlexibarBatch.ParentFolder)
                End If

                LogPostEvaluationScript(lobjImagen)

            End If

            pstrExportMessage = ""
            Return True

        Catch ex As Exception

            pstrExportMessage = ex.Message.ToString()
            Return False
        End Try

    End Function

    ' *******************************************************************************************************
    ' LogPostEvaluationScript
    ' Desc: Rutina que escribe en el log lo que se ha leido
    ' NBL 12/11/2009
    ' *******************************************************************************************************
    Private Sub LogPostEvaluationScript(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lstrRutaLog As String = UtilGlobal.UShared.GetFolderLaboLog

        If Not My.Computer.FileSystem.DirectoryExists(lstrRutaLog) Then Exit Sub

        Dim lstrFileLog As String = lstrRutaLog & "\" & Now.Year & Microsoft.VisualBasic.Right("00" & Now.Month, 2) & _
                                                                                                                    Microsoft.VisualBasic.Right("00" & Now.Day, 2) & ".txt"

        Dim lobjStreamWriter As New IO.StreamWriter(lstrFileLog, True)

        lobjStreamWriter.WriteLine("")
        lobjStreamWriter.WriteLine("")
        lobjStreamWriter.WriteLine("--------------------------------------------------------------------------------------------------------")
        lobjStreamWriter.WriteLine("--------------------------------------------------------------------------------------------------------")

        lobjStreamWriter.WriteLine("DNoPet: " & pobjImage.VirtualFields.GetFieldValue("DNoPet"))
        lobjStreamWriter.WriteLine("DNoPet2: " & pobjImage.VirtualFields.GetFieldValue("DNoPet2"))
        lobjStreamWriter.WriteLine("DNoPet3: " & pobjImage.VirtualFields.GetFieldValue("DNoPet3"))
        lobjStreamWriter.WriteLine("DNoHist: " & pobjImage.VirtualFields.GetFieldValue("DNoHist"))
        lobjStreamWriter.WriteLine("DNoHist2: " & pobjImage.VirtualFields.GetFieldValue("DNoHist2"))
        lobjStreamWriter.WriteLine("DNoHist3: " & pobjImage.VirtualFields.GetFieldValue("DNoHist3"))
        lobjStreamWriter.WriteLine("DNoHistFusion: " & pobjImage.VirtualFields.GetFieldValue("DNoHistFusion"))
        lobjStreamWriter.WriteLine("DEpisodio: " & pobjImage.VirtualFields.GetFieldValue("DEpisodio"))
        lobjStreamWriter.WriteLine("DActo: " & pobjImage.VirtualFields.GetFieldValue("DActo"))
        lobjStreamWriter.WriteLine("DNoSS: " & pobjImage.VirtualFields.GetFieldValue("DNoSS"))
        lobjStreamWriter.WriteLine("DDNI: " & pobjImage.VirtualFields.GetFieldValue("DDNI"))
        lobjStreamWriter.WriteLine("DApellido1: " & pobjImage.VirtualFields.GetFieldValue("DApellido1"))
        lobjStreamWriter.WriteLine("DApellido2: " & pobjImage.VirtualFields.GetFieldValue("DApellido2"))
        lobjStreamWriter.WriteLine("DNombre: " & pobjImage.VirtualFields.GetFieldValue("DNombre"))
        lobjStreamWriter.WriteLine("DFechaNac: " & pobjImage.VirtualFields.GetFieldValue("DFechaNac"))
        lobjStreamWriter.WriteLine("DSexo: " & pobjImage.VirtualFields.GetFieldValue("DSexo"))
        lobjStreamWriter.WriteLine("DDireccion: " & pobjImage.VirtualFields.GetFieldValue("DDireccion"))
        lobjStreamWriter.WriteLine("DTelefono: " & pobjImage.VirtualFields.GetFieldValue("DTelefono"))
        lobjStreamWriter.WriteLine("DPoblacion: " & pobjImage.VirtualFields.GetFieldValue("DPoblacion"))
        lobjStreamWriter.WriteLine("DCPostal: " & pobjImage.VirtualFields.GetFieldValue("DCPostal"))
        lobjStreamWriter.WriteLine("DDoctor: " & pobjImage.VirtualFields.GetFieldValue("DDoctor"))
        lobjStreamWriter.WriteLine("DTDoctor: " & pobjImage.VirtualFields.GetFieldValue("DTDoctor"))
        lobjStreamWriter.WriteLine("DFactura: " & pobjImage.VirtualFields.GetFieldValue("DFactura"))
        lobjStreamWriter.WriteLine("DCDiagnostico: " & pobjImage.VirtualFields.GetFieldValue("DCDiagnostico"))
        lobjStreamWriter.WriteLine("DTDiagnostico: " & pobjImage.VirtualFields.GetFieldValue("DTDiagnostico"))
        lobjStreamWriter.WriteLine("DPrioridad: " & pobjImage.VirtualFields.GetFieldValue("DPrioridad"))
        lobjStreamWriter.WriteLine("DCama: " & pobjImage.VirtualFields.GetFieldValue("DCama"))
        lobjStreamWriter.WriteLine("DTipo: " & pobjImage.VirtualFields.GetFieldValue("DTipo"))
        lobjStreamWriter.WriteLine("DMotivo: " & pobjImage.VirtualFields.GetFieldValue("DMotivo"))
        lobjStreamWriter.WriteLine("DServicio: " & pobjImage.VirtualFields.GetFieldValue("DServicio"))
        lobjStreamWriter.WriteLine("DOrigen: " & pobjImage.VirtualFields.GetFieldValue("DOrigen"))
        lobjStreamWriter.WriteLine("DDestino: " & pobjImage.VirtualFields.GetFieldValue("DDestino"))
        lobjStreamWriter.WriteLine("DGrupo: " & pobjImage.VirtualFields.GetFieldValue("DGrupo"))
        lobjStreamWriter.WriteLine("DTipoFisiologico: " & pobjImage.VirtualFields.GetFieldValue("DTipoFisiologico"))
        lobjStreamWriter.WriteLine("DFormID: " & pobjImage.VirtualFields.GetFieldValue("DFormID"))
        lobjStreamWriter.WriteLine("DObservaciones: " & pobjImage.VirtualFields.GetFieldValue("DObservaciones"))
        lobjStreamWriter.WriteLine("DFHExtraccion: " & pobjImage.VirtualFields.GetFieldValue("DFHExtraccion"))
        lobjStreamWriter.WriteLine("DFHRegistro: " & pobjImage.VirtualFields.GetFieldValue("DFHRegistro"))
        lobjStreamWriter.WriteLine("DPruebas: " & pobjImage.VirtualFields.GetFieldValue("DPruebas"))
        lobjStreamWriter.WriteLine("DPerfiles: " & pobjImage.VirtualFields.GetFieldValue("DPerfiles"))
        lobjStreamWriter.WriteLine("DMuestra: " & pobjImage.VirtualFields.GetFieldValue("DMuestra"))
        lobjStreamWriter.WriteLine("DResultados: " & pobjImage.VirtualFields.GetFieldValue("DResultados"))
        lobjStreamWriter.WriteLine("DNTelefono: " & pobjImage.VirtualFields.GetFieldValue("DNTelefono"))
        lobjStreamWriter.WriteLine("DFaxResultados: " & pobjImage.VirtualFields.GetFieldValue("DFaxResultados"))
        lobjStreamWriter.WriteLine("DScanStation: " & pobjImage.VirtualFields.GetFieldValue("DScanStation"))
        lobjStreamWriter.WriteLine("DBatchNo: " & pobjImage.VirtualFields.GetFieldValue("DBatchNo"))
        lobjStreamWriter.WriteLine("DPageNo: " & pobjImage.VirtualFields.GetFieldValue("DPageNo"))
        lobjStreamWriter.WriteLine("DUser: " & pobjImage.VirtualFields.GetFieldValue("DUser"))
        lobjStreamWriter.WriteLine("TOrigen: " & pobjImage.VirtualFields.GetFieldValue("TOrigen"))
        lobjStreamWriter.WriteLine("TDestino: " & pobjImage.VirtualFields.GetFieldValue("TDestino"))
        lobjStreamWriter.WriteLine("TServicio: " & pobjImage.VirtualFields.GetFieldValue("TServicio"))
        lobjStreamWriter.WriteLine("AOrina24: " & pobjImage.VirtualFields.GetFieldValue("AOrina24"))
        lobjStreamWriter.WriteLine("ACodigosMarcas: " & pobjImage.VirtualFields.GetFieldValue("ACodigosMarcas"))
        lobjStreamWriter.WriteLine("ACodigosManuscrito: " & pobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito"))

        lobjStreamWriter.Close()

    End Sub


    ' **********************************************************************************************
    ' CompruebaExportacionParticularPlantilla
    ' Desc.: Rutina que comprueba si existe una exportación particular para una plantilla en concreto
    ' NBL 7/10/2009
    ' **********************************************************************************************
    Private Sub CompruebaExportacionParticularPlantilla(ByRef pobjImagen As LibFlexibarNETObjects.Image)

        If pobjImagen.ARData.TemplateName Is Nothing Then
            Me.mbolExisteConfiguracionParticular = False
            Exit Sub
        End If

        If pobjImagen.ARData.TemplateName.Length = 0 Then
            Me.mbolExisteConfiguracionParticular = False
            Exit Sub
        End If

        If mobjINI.IniGet(mstrRutaINI, pobjImagen.ARData.TemplateName, "Activada", 0) = 1 Then
            Me.mbolExisteConfiguracionParticular = True
            ' Como hay una exportación propia para la plantilla en cuestion, la cargamos
            CargarConfiguracionParticular(pobjImagen.ARData.TemplateName)
            CargarConfiguracionParticularBackup(pobjImagen.ARData.TemplateName)
        Else
            Me.mbolExisteConfiguracionParticular = False
        End If

    End Sub

    ' **********************************************************************************************
    ' ExportarArchivoASTM
    ' Desc.: Rutina que exporta el archivo que contiene la cadena ASTM
    ' NBL 8/10/2009
    ' **********************************************************************************************
    Private Sub ExportarArchivoASTM(ByRef pobjImagen As LibFlexibarNETObjects.Image, _
                                                                ByVal pstrRutaBase As String)

        If pstrRutaBase.ToUpper.IndexOf(",") > -1 Then

            Dim objFTP As New FTP
            Dim strFTP() As String = pstrRutaBase.Split(",")

            Dim lstrRutaDestino As String = pobjImagen.VirtualFields.GetFieldValue("DNoPet")
            Dim lstrRutaDestino2 As String = lstrRutaDestino

            Dim lintContador As Integer = 1
            '' Miramos si existe o no
            If objFTP.ExistFile(lstrRutaDestino & "." & Me.mstrExtensionASTM, strFTP(0), strFTP(1), strFTP(2), strFTP(3)) Then

                Do
                    lstrRutaDestino = lstrRutaDestino2 & "-" & lintContador.ToString

                    If Not objFTP.ExistFile(lstrRutaDestino & "." & Me.mstrExtensionASTM, strFTP(0), strFTP(1), strFTP(2), strFTP(3)) Then Exit Do
                    lintContador += 1
                Loop
            End If


            'objFTP.SendText(Me.mstrCadenaFinal, strFTP(0), strFTP(1), strFTP(2), lstrRutaDestino & "." & Me.mstrExtensionASTM, strFTP(3))
            objFTP.SendText(Me.mstrCadenaFinal, strFTP(0), strFTP(1), strFTP(2), lstrRutaDestino & ".tmp", strFTP(3))
            objFTP.RenameRemoteFile(lstrRutaDestino & ".tmp", lstrRutaDestino & "." & Me.mstrExtensionASTM, strFTP(0), strFTP(1), strFTP(2), strFTP(3))

        Else
            ' Primero tenemos que exportar el archivo de texto con la cadena ASTM al lugar de proceso
            ' después hacemos lo mismo si hay que enviarlo  a la ruta  de backup
            If Not My.Computer.FileSystem.DirectoryExists(pstrRutaBase) Then _
                                                             Throw New Exception("La carpeta no existe: " & pstrRutaBase)

            Dim lstrRutaDestino As String = pstrRutaBase & "\" & pobjImagen.VirtualFields.GetFieldValue("DNoPet") & "." & Me.mstrExtensionASTM

            ' Miramos si existe o no
            If My.Computer.FileSystem.FileExists(lstrRutaDestino) Then
                Dim lintContador As Integer = 1
                Do
                    lstrRutaDestino = lstrRutaDestino.Substring(0, lstrRutaDestino.Length - 4) & "-" & lintContador.ToString & _
                                            "." & Me.mstrExtensionASTM
                    If Not My.Computer.FileSystem.FileExists(lstrRutaDestino) Then Exit Do
                    lintContador += 1
                Loop
            End If

            Dim lobjStreamWriter As New StreamWriter(lstrRutaDestino, True, System.Text.Encoding.GetEncoding(1252))
            lobjStreamWriter.Write(Me.mstrCadenaFinal)
            lobjStreamWriter.Close()
        End If
    End Sub

    ' **********************************************************************************************
    ' ExportarArchivoASTM
    ' Desc.: Rutina que exporta el archivo que contiene la cadena ASTM
    ' NBL 8/10/2009
    ' **********************************************************************************************
    Private Sub ExportarArchivoASTM(ByRef pobjImagen As LibFlexibarNETObjects.Image, _
                                                                ByVal pstrRutaBase As String, ByVal pdtFechaLote As Date)

        ' Primero tenemos que exportar el archivo de texto con la cadena ASTM al lugar de proceso
        ' después hacemos lo mismo si hay que enviarlo a la ruta de backup
        If Not My.Computer.FileSystem.DirectoryExists(pstrRutaBase) Then _
                                                         Throw New Exception("La carpeta no existe: " & pstrRutaBase)

        Dim lstrCarpetaDestino As String = pstrRutaBase & "\" & pdtFechaLote.Year & Microsoft.VisualBasic.Right("00" & pdtFechaLote.Month, 2) & _
                                   Microsoft.VisualBasic.Right("00" & pdtFechaLote.Day, 2)

        If Not My.Computer.FileSystem.DirectoryExists(lstrCarpetaDestino) Then My.Computer.FileSystem.CreateDirectory(lstrCarpetaDestino)

        Dim lstrRutaDestino As String = lstrCarpetaDestino & "\" & pobjImagen.VirtualFields.GetFieldValue("DNoPet") & "." & Me.mstrExtensionASTM
        'Dim lstrRutaDestino As String = lstrCarpetaDestino & "\" & mstrNoPet & "." & Me.mstrExtensionASTM

        ' Miramos si existe o no
        If My.Computer.FileSystem.FileExists(lstrRutaDestino) Then
            Dim lintContador As Integer = 1
            Do
                lstrRutaDestino = lstrRutaDestino.Substring(0, lstrRutaDestino.Length - 4) & "-" & lintContador.ToString & _
                                        "." & Me.mstrExtensionASTM
                If Not My.Computer.FileSystem.FileExists(lstrRutaDestino) Then Exit Do
                lintContador += 1
            Loop
        End If

        Dim lobjStreamWriter As New StreamWriter(lstrRutaDestino)
        lobjStreamWriter.Write(Me.mstrCadenaFinal)
        lobjStreamWriter.Close()

    End Sub

    ' **********************************************************************************************
    ' ExportarImagen
    ' Desc.: Rutina que exporta la imagen a a destino
    ' NBL 7/10/2009
    ' **********************************************************************************************
    Private Sub ExportarImagen(ByRef pobjImagen As LibFlexibarNETObjects.Image, ByVal pstrRutaBase As String,
                                                    ByVal pbolSubdirectorioFecha As Boolean, ByVal pintFechaSubcarpeta As Integer,
                                                    ByVal pbolAnexarImagenes As Boolean, ByVal pdtFechaLote As Date, ByVal pstrParentFolder As String)

        If Not My.Computer.FileSystem.DirectoryExists(pstrRutaBase) Then Throw New Exception("No existe la carpeta: " & pstrRutaBase)

        Dim lstrCarpetaDestino As String = pstrRutaBase

        If pbolSubdirectorioFecha Then
            If pintFechaSubcarpeta = 1 And pobjImagen.VirtualFields.GetFieldValue("DFHExtraccion").Trim.Length = 8 Then
                lstrCarpetaDestino &= "\" & pobjImagen.VirtualFields.GetFieldValue("DFHExtraccion")
            ElseIf pintFechaSubcarpeta = 2 And pobjImagen.VirtualFields.GetFieldValue("DFHRegistro").Trim.Length = 8 Then
                lstrCarpetaDestino &= "\" & pobjImagen.VirtualFields.GetFieldValue("DFHRegistro")
            ElseIf pintFechaSubcarpeta = 2 Then
                lstrCarpetaDestino &= "\" & pdtFechaLote.Year & Microsoft.VisualBasic.Right("00" & pdtFechaLote.Month, 2) &
                                    Microsoft.VisualBasic.Right("00" & pdtFechaLote.Day, 2)
            End If
            If Not My.Computer.FileSystem.DirectoryExists(lstrCarpetaDestino) Then _
                            My.Computer.FileSystem.CreateDirectory(lstrCarpetaDestino)
        End If

        Dim lstrRutaImagen As String = ""

        If Me.mbolNombreImagenMMDD Then
            lstrRutaImagen = lstrCarpetaDestino & "\" & Now.Month & Microsoft.VisualBasic.Right("00" & Now.Day, 2) & pobjImagen.VirtualFields.GetFieldValue("DNoPet") & ".tif"
        Else
            lstrRutaImagen = lstrCarpetaDestino & "\" & pobjImagen.VirtualFields.GetFieldValue("DNoPet") & ".tif"
        End If

        If Not My.Computer.FileSystem.FileExists(lstrRutaImagen) Then
            My.Computer.FileSystem.CopyFile(pobjImagen.Path(pstrParentFolder), lstrRutaImagen)
        ElseIf My.Computer.FileSystem.FileExists(lstrRutaImagen) And pbolAnexarImagenes = True Then
            mobjImage.UnirTiFFs(lstrRutaImagen, pobjImagen.Path(pstrParentFolder))
        Else
            ' Aquí hacemos un bucle mirando el nombre añadiendo un contador al final del nombre de la imagen
            Dim lintContador As Integer = 1
            Do
                Dim lstrRutaContador As Integer = lstrRutaImagen.Substring(0, lstrRutaImagen.Length - 4) & "-" & lintContador.ToString() & ".tif"
                If Not My.Computer.FileSystem.FileExists(lstrRutaContador) Then
                    My.Computer.FileSystem.CopyFile(pobjImagen.Path(pstrParentFolder), lstrRutaContador)
                    Exit Do
                End If
                lintContador += 1
            Loop
        End If

    End Sub

    ' **********************************************************************************************
    ' InicializarExportASTM
    ' Desc.: Rutina que carga la configuracion ASTM del archivo ino
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub InicializarExportASTM()

        CargarConfiguracionGeneral()
        CargarConfiguracionStandard()
        CargarConfiguracionStandardBackup()

    End Sub

    ' **********************************************************************************************
    ' CargarConfiguracionStandard
    ' Desc.: Cargamos la parte de configuración que es standard
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub CargarConfiguracionStandard()

        ' ACTIVADA
        If mobjINI.IniGet(mstrRutaINI, "Standard", "Activada", "0") = "1" Then
            Me.mbolStdActivada = True
        Else
            Me.mbolStdActivada = False
        End If
        ' RUTA BASE EXPORTACIÓN
        Me.mstrStdRutaBaseExportacion = mobjINI.IniGet(mstrRutaINI, "Standard", "RutaBaseExportacion", "")
        ' ACTIVAR EXPORTACION IMAGEN
        If mobjINI.IniGet(mstrRutaINI, "Standard", "ActivarExportacionImagen", "0") = "1" Then
            Me.mbolStdActivarExportacionImagen = True
        Else
            Me.mbolStdActivarExportacionImagen = False
        End If
        ' RUTABASE IMAGEN
        Me.mstrStdRutaBaseImagen = mobjINI.IniGet(mstrRutaINI, "Standard", "RutaBaseImagen", "")
        ' SUBDIRECTORIOFECHA
        If mobjINI.IniGet(mstrRutaINI, "Standard", "SubdirectorioFecha", "0") = "1" Then
            Me.mbolStdSubdirectorioFecha = True
        Else
            Me.mbolStdSubdirectorioFecha = False
        End If
        ' FECHA
        Me.mintStdFecha = mobjINI.IniGet(mstrRutaINI, "Standard", "Fecha", 0)
        ' ANEXAR IMÁGENES
        If mobjINI.IniGet(mstrRutaINI, "Standard", "AnexarImagenes", "0") = "1" Then
            Me.mbolStdAnexarImagenes = True
        Else
            Me.mbolStdAnexarImagenes = False
        End If

        ' 20/09/2010 PVT: Cargamos la asignación de fechas y horas de extracción y de Registro.
        If mobjINI.IniGet(mstrRutaINI, "Standard", "BatchDate2FRegistro", "0") = "1" Then
            Me.mbolStdBatchDateToFRegistro = True
        Else
            Me.mbolStdBatchDateToFRegistro = False
        End If

        If mobjINI.IniGet(mstrRutaINI, "Standard", "BatchDate2FExtraccion", "0") = "1" Then
            Me.mbolStdBatchDateToFExtraccion = True
        Else
            Me.mbolStdBatchDateToFExtraccion = False
        End If

        If mobjINI.IniGet(mstrRutaINI, "Standard", "AddHora2FRegistro", "0") = "1" Then
            Me.mbolStdAddHoraToFRegistro = True
        Else
            Me.mbolStdAddHoraToFRegistro = False
        End If

        If mobjINI.IniGet(mstrRutaINI, "Standard", "AddHora2FExtraccion", "0") = "1" Then
            Me.mbolStdAddHoraToFExtraccion = True
        Else
            Me.mbolStdAddHoraToFExtraccion = False
        End If


    End Sub

    ' **********************************************************************************************
    ' CargarConfiguracionParticular
    ' Desc.: Cargamos la parte de configuración que es particular para la plantilla de la imagen
    ' NBL 7/10/2009
    ' **********************************************************************************************
    Private Sub CargarConfiguracionParticular(ByVal pstrPlantilla As String)

        ' ACTIVADA
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "Activada", "0") = "1" Then
            Me.mbolPrtActivada = True
        Else
            Me.mbolPrtActivada = False
        End If
        ' RUTA BASE EXPORTACIÓN
        Me.mstrPrtRutaBaseExportacion = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "RutaBaseExportacion", "")
        ' ACTIVAR EXPORTACION IMAGEN
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "ActivarExportacionImagen", "0") = "1" Then
            Me.mbolPrtActivarExportacionImagen = True
        Else
            Me.mbolPrtActivarExportacionImagen = False
        End If
        ' RUTABASE IMAGEN
        Me.mstrPrtRutaBaseImagen = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "RutaBaseImagen", "")
        ' SUBDIRECTORIOFECHA
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "SubdirectorioFecha", "0") = "1" Then
            Me.mbolPrtSubdirectorioFecha = True
        Else
            Me.mbolPrtSubdirectorioFecha = False
        End If
        ' FECHA
        Me.mintPrtFecha = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "Fecha", 0)
        ' ANEXAR IMÁGENES
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "AnexarImagenes", "0") = "1" Then
            Me.mbolPrtAnexarImagenes = True
        Else
            Me.mbolPrtAnexarImagenes = False
        End If

        ' 20/09/2010 PVT: Cargamos la asignación de fechas y horas de extracción y de Registro.

        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "BatchDate2FRegistro", "0") = "1" Then
            Me.mbolPrtBatchDateToFRegistro = True
        Else
            Me.mbolPrtBatchDateToFRegistro = False
        End If

        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "BatchDate2FExtraccion", "0") = "1" Then
            Me.mbolPrtBatchDateToFExtraccion = True
        Else
            Me.mbolPrtBatchDateToFExtraccion = False
        End If

        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "AddHora2FRegistro", "0") = "1" Then
            Me.mbolPrtAddHoraToFRegistro = True
        Else
            Me.mbolPrtAddHoraToFRegistro = False
        End If

        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "AddHora2FExtraccion", "0") = "1" Then
            Me.mbolPrtAddHoraToFExtraccion = True
        Else
            Me.mbolPrtAddHoraToFExtraccion = False
        End If

    End Sub

    ' **********************************************************************************************
    ' CargarConfiguracionStandardBackup
    ' Desc.: Cargamos la parte de configuración que es standard
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub CargarConfiguracionStandardBackup()

        ' ACTIVADA
        If mobjINI.IniGet(mstrRutaINI, "Standard", "ActivadaBackup", "0") = "1" Then
            Me.mbolStdActivadaBackup = True
        Else
            Me.mbolStdActivadaBackup = False
        End If
        ' RUTA BASE EXPORTACIÓN
        Me.mstrStdRutaBaseExportacionBackup = mobjINI.IniGet(mstrRutaINI, "Standard", "RutaBaseExportacionBackup", "")
        ' ACTIVAR EXPORTACION IMAGEN
        If mobjINI.IniGet(mstrRutaINI, "Standard", "ActivarExportacionImagenBackup", "0") = "1" Then
            Me.mbolStdActivarExportacionImagenBackup = True
        Else
            Me.mbolStdActivarExportacionImagenBackup = False
        End If
        ' RUTABASE IMAGEN
        Me.mstrStdRutaBaseImagenBackup = mobjINI.IniGet(mstrRutaINI, "Standard", "RutaBaseImagenBackup", "")
        ' SUBDIRECTORIOFECHA
        If mobjINI.IniGet(mstrRutaINI, "Standard", "SubdirectorioFechaBackup", "0") = "1" Then
            Me.mbolStdSubdirectorioFechaBackup = True
        Else
            Me.mbolStdSubdirectorioFechaBackup = False
        End If
        ' FECHA
        Me.mintStdFechaBackup = mobjINI.IniGet(mstrRutaINI, "Standard", "FechaBackup", 0)
        ' ANEXAR IMÁGENES
        If mobjINI.IniGet(mstrRutaINI, "Standard", "AnexarImagenesBackup", "0") = "1" Then
            Me.mbolStdAnexarImagenesBackup = True
        Else
            Me.mbolStdAnexarImagenesBackup = False
        End If

    End Sub

    ' **********************************************************************************************
    ' CargarConfiguracionParticularBackup
    ' Desc.: Cargamos la parte de configuración que es de plantilla backup
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub CargarConfiguracionParticularBackup(ByVal pstrPlantilla As String)

        ' ACTIVADA
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "ActivadaBackup", "0") = "1" Then
            Me.mbolPrtActivadaBackup = True
        Else
            Me.mbolPrtActivadaBackup = False
        End If
        ' RUTA BASE EXPORTACIÓN
        Me.mstrPrtRutaBaseExportacionBackup = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "RutaBaseExportacionBackup", "")
        ' ACTIVAR EXPORTACION IMAGEN
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "ActivarExportacionImagenBackup", "0") = "1" Then
            Me.mbolPrtActivarExportacionImagenBackup = True
        Else
            Me.mbolPrtActivarExportacionImagenBackup = False
        End If
        ' RUTABASE IMAGEN
        Me.mstrPrtRutaBaseImagenBackup = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "RutaBaseImagenBackup", "")
        ' SUBDIRECTORIOFECHA
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "SubdirectorioFechaBackup", "0") = "1" Then
            Me.mbolPrtSubdirectorioFechaBackup = True
        Else
            Me.mbolPrtSubdirectorioFechaBackup = False
        End If
        ' FECHA
        Me.mintPrtFechaBackup = mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "FechaBackup", 0)
        ' ANEXAR IMÁGENES
        If mobjINI.IniGet(mstrRutaINI, pstrPlantilla, "AnexarImagenesBackup", "0") = "1" Then
            Me.mbolPrtAnexarImagenesBackup = True
        Else
            Me.mbolPrtAnexarImagenesBackup = False
        End If

    End Sub

    ' **********************************************************************************************
    ' CargarConfiguracionGeneral
    ' Desc.: Cargamos la parte de configuración que es general
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub CargarConfiguracionGeneral()

        ' ACTIVAR EXPORTACION ASTM
        If mobjINI.IniGet(mstrRutaINI, "General", "ActivarExportacionASTM", "0") = "1" Then
            Me.mbolActivarExportacionASTM = True
        Else
            Me.mbolActivarExportacionASTM = False
        End If
        ' EXTENSIÓN ASTM
        Me.mstrExtensionASTM = mobjINI.IniGet(mstrRutaINI, "General", "ExtensionASTM", "")
        ' ACTIVAR DEBUG
        If mobjINI.IniGet(mstrRutaINI, "General", "ActivarDebug", "0") = "1" Then
            Me.mbolActivarDebug = True
        Else
            Me.mbolActivarDebug = False
        End If
        ' RUTA BASE DEBUG
        Me.mstrRutaBaseDebug = mobjINI.IniGet(mstrRutaINI, "General", "RutaBaseDebug", "")
        ' INTERACTIVA
        If mobjINI.IniGet(mstrRutaINI, "General", "Interactiva", "0") = "1" Then
            Me.mbolInteractiva = True
        Else
            Me.mbolInteractiva = False
        End If
        ' MMDD
        If mobjINI.IniGet(mstrRutaINI, "General", "MesDia", "0") = "1" Then
            Me.mbolMMDD = True
        Else
            Me.mbolMMDD = False
        End If

        ' NombreImagen
        If mobjINI.IniGet(mstrRutaINI, "General", "MesDiaNombreImagen", "0") = "1" Then
            Me.mbolNombreImagenMMDD = True
        Else
            Me.mbolNombreImagenMMDD = False
        End If

    End Sub

    ' **********************************************************************************************
    ' ModificarNoPeticion
    ' Desc.: Rutina que pone el numero el mes y el dia delante del número de petición
    ' NBL 20/11/2009
    ' **********************************************************************************************
    Private Sub ModificarNoPeticion(ByRef pobjImage As LibFlexibarNETObjects.Image)

        If Me.mbolMMDD Then
            'pobjImage.VirtualFields.SetFieldValue("DNoPet", Microsoft.VisualBasic.Right("00" & Now.Month, 2) & Microsoft.VisualBasic.Right("00" & Now.Day, 2) & pobjImage.VirtualFields.GetFieldValue("DNoPet"))
            mstrNoPet = Now.Month & Microsoft.VisualBasic.Right("00" & Now.Day, 2) & pobjImage.VirtualFields.GetFieldValue("DNoPet")
        Else
            mstrNoPet = pobjImage.VirtualFields.GetFieldValue("DNoPet")
        End If

    End Sub

    ' **********************************************************************************************
    ' CreateStringASTM
    ' Desc.: Rutina principal de creación de la cadena ASTM
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub CreateStringASTM(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pdtBatchDate As Date)

        ModificarNoPeticion(pobjImage)

        InformacionCabecera()
        InformacionDelPaciente(pobjImage, pdtBatchDate)
        ' Creación de las líneas de pruebas
        Dim pintNumeroOrden As Integer = 0
        LineasCodigosPrueba(pobjImage, pintNumeroOrden)
        LineasPxxxxA_PxxxxR(pobjImage, pintNumeroOrden)
        ' OJO, falta indicar las líneas de comentario

        FinCreacionCadenaASTM()

    End Sub

    ' **********************************************************************************************
    ' LineasPxxxxA_PxxxxR
    ' Desc.: Rutina que trata este tipo de variables virtuales
    ' NBL 17/12/2009
    ' **********************************************************************************************
    Private Sub LineasPxxxxA_PxxxxR(ByRef pobjImage As LibFlexibarNETObjects.Image, ByRef pintNumeroOrden As Integer)

        For Each lobjVirtualField As LibFlexibarNETObjects.VirtualField In pobjImage.VirtualFields

            If Regex.IsMatch(lobjVirtualField.Name, "^P\d{4}A") Then

                pintNumeroOrden += 1

                Me.mstrCadenaFinal &= "O|" & pintNumeroOrden & "|"
                'Me.mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(pobjImage.VirtualFields.GetFieldValue("DNoPet")) & "||"
                Me.mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(mstrNoPet) & "||"
                Me.mstrCadenaFinal &= lobjVirtualField.Text & vbCrLf

                Dim lstrNombreR As String = lobjVirtualField.Name.Replace("A", "R")
                Dim lstrValor As String = pobjImage.VirtualFields.GetFieldValue(lstrNombreR)

                Me.mstrCadenaFinal &= "R|1|"
                Me.mstrCadenaFinal &= lstrValor & vbCrLf

            End If

        Next

    End Sub
    ' **********************************************************************************************
    ' LineasCodigosPrueba
    ' Desc.: Rutina que procesa los códigos de pruebas que tiene la imagen para crear las lineas ASTM correspondientes
    ' NBL 5/10/2009
    ' **********************************************************************************************
    Private Sub LineasCodigosPrueba(ByRef pobjImage As LibFlexibarNETObjects.Image, ByRef pintNumeroOrden As Integer)

        Dim lstrCodigosMarcas As String = pobjImage.VirtualFields.GetFieldValue("ACodigosMarcas")
        Dim lstrCodigosManuscritos As String = pobjImage.VirtualFields.GetFieldValue("ACodigosManuscrito")

        Dim lstrCodigos As String = lstrCodigosMarcas & lstrCodigosManuscritos

        ' No hay nada
        If lstrCodigos.Trim.Length = 0 Then Exit Sub

        Dim lstrParseaCodigos() As String = lstrCodigos.Split(",")
        Dim lstrTipoMuestraGeneral As String = lstrParseaCodigos(0).ToString()

        If lstrParseaCodigos.Length > 1 Then
            ' Bucle por cada una de las pruebas

            For lintContador As Integer = 1 To lstrParseaCodigos.Length - 1
                Dim lstrPrueba As String = lstrParseaCodigos(lintContador).ToString()

                If lstrPrueba.IndexOf("|") <> -1 Then
                    'TratamientoPruebasEmpaquetadas(lstrTipoMuestraGeneral, lstrPrueba, lintContador, pobjImage.VirtualFields.GetFieldValue("DNoPet"))
                    TratamientoPruebasEmpaquetadas(lstrTipoMuestraGeneral, lstrPrueba, lintContador, mstrNoPet)
                Else
                    'TratamientoNOEmpaquetadas(lstrTipoMuestraGeneral, lstrPrueba, lintContador, pobjImage.VirtualFields.GetFieldValue("DNoPet"))
                    TratamientoNOEmpaquetadas(lstrTipoMuestraGeneral, lstrPrueba, lintContador, mstrNoPet)
                End If

            Next

            pintNumeroOrden = lstrParseaCodigos.Length - 1

        End If

    End Sub

    ' **********************************************************************************************
    ' TratamientoPruebasEmpaquetadas
    ' Desc.: 
    ' NBL 6/10/2009
    ' **********************************************************************************************
    Private Sub TratamientoPruebasEmpaquetadas(ByVal pstrMuestraDefecto As String, ByVal pstrPrueba As String, _
                                                                                ByVal pintContadorPrueba As Integer, ByVal pstrNumeroPeticion As String)

        Dim lstrPrueba() As String = pstrPrueba.Split("|")
        Dim lstrP1 As String = "", lstrP2 As String = "", lstrP3 As String = ""
        lstrP1 = lstrPrueba(0)
        If lstrPrueba.Length > 1 Then lstrP2 = lstrPrueba(1)
        If lstrPrueba.Length > 2 Then lstrP3 = lstrPrueba(2)

        Me.mstrCadenaFinal &= "O|" & pintContadorPrueba & "|" 'Numero de orden
        ' Número de petición
        Me.mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(pstrNumeroPeticion) & "||"
        ' Código de la prueba
        If lstrP1.IndexOf("^") <> -1 Then
            ' Hay tipo de prueba
            Me.mstrCadenaFinal &= "^^^" & lstrP1 & "^^|||||||A||||" 'Ver que me he comido el ^B de bioquimica
        Else
            ' No hay tipo de prueba
            Me.mstrCadenaFinal &= "^^^" & lstrP1 & "^" & mstrTipoPruebaPorDefecto & "^^|||||||A||||"
        End If

        ' Tipo de muestra
        If lstrP2 = "" Then ' Valor por defecto, el general M0
            Me.mstrCadenaFinal &= pstrMuestraDefecto & "|||"
        Else
            Me.mstrCadenaFinal &= lstrP2 & "|||"
        End If

        ' Seccion
        Me.mstrCadenaFinal &= lstrP3 & "|||||||O" & vbCrLf

    End Sub

    ' *********************************************************************************
    ' TratamientoNOEmpaquetadas
    ' Desc:
    ' NBL: 06/10/2009
    ' *********************************************************************************
    Private Sub TratamientoNOEmpaquetadas(ByVal pstrMuestraDefecto As String, ByVal pstrPrueba As String, _
                                                                                ByVal pintContadorPrueba As Integer, ByVal pstrNumeroPeticion As String)

        ' Sin empaquetar
        Me.mstrCadenaFinal &= "O|" & pintContadorPrueba & "|" ' Número de orden
        ' Número de petición
        Me.mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(pstrNumeroPeticion) & "||"
        ' Código de prueba

        If pstrPrueba.IndexOf("^") Then
            ' Hay tipo de prueba
            Me.mstrCadenaFinal &= "^^^" & pstrPrueba & "^^|||||||A||||"
        Else
            'No hay tipo de prueba (Default en TdP)
            Me.mstrCadenaFinal &= "^^^" & pstrPrueba & "^" & mstrTipoPruebaPorDefecto & "^^|||||||A||||"
        End If

        '16 Tipo de muestra. OJO NO EXISTE NUNCA P####M, siempre es el valor por defecto
        Me.mstrCadenaFinal &= pstrMuestraDefecto & "|||"

        '19 Seccion. OJO Existencia NO EXISTE P####S , siempre vacio
        Me.mstrCadenaFinal &= "|||||||O" & vbCrLf

    End Sub



    ' **********************************************************************************************
    ' InformacionCabecera
    ' Desc.: Función que escribe la primera línea del ASTM
    ' NBL 17/7/2009
    ' **********************************************************************************************
    Private Sub InformacionCabecera()

        mstrCadenaFinal &= "H|\^&|||Tecnomedia|SCANU||||OMEGA||P||" & Format$(Now, "yyyyMMddHHmmss") & vbCrLf

    End Sub

    ' ***********************************************************************************************
    ' FinCreacionCadenaASTM
    ' Desc.: Fin de la cadena
    ' NBL 17/7/2009
    ' ***********************************************************************************************
    Private Sub FinCreacionCadenaASTM()

        mstrCadenaFinal &= "L|1|F" & vbCrLf

    End Sub

    ' ***********************************************************************************************
    ' FormateaFechaNac
    ' Desc.: Formateamos la fecha de nacimiento de forma correcta
    ' NBL 7/10/2010
    ' ***********************************************************************************************
    Private Function FormateaFechaNac(ByVal pstrFecha As String) As String

        If pstrFecha.Trim.Length = 0 Then Return ""

        If pstrFecha.Length = 8 Then Return pstrFecha

        If pstrFecha.Length = 10 Then
            Dim lstrResultado As String = pstrFecha.Substring(6, 4) & pstrFecha.Substring(3, 2) & pstrFecha.Substring(0, 2)
            Return lstrResultado
        End If

        Return ""

    End Function

    ' ***********************************************************************************************
    ' InformacionDelPaciente
    ' Desc.: Añadimos la línea de pacientes
    ' NBL 17/7/2009
    ' ***********************************************************************************************
    Private Sub InformacionDelPaciente(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pdtBatchDate As Date)

        mstrCadenaFinal &= "P|1|"
        ' Número de petición
        'mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(pobjImage.VirtualFields.GetFieldValue("DNoPet").Trim) & "|"
        mstrCadenaFinal &= UtilGlobal.UShared.QuitarCerosIzquierda(mstrNoPet.Trim) & "|"
        ' Número de historia
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DNoHist").Trim & "|"
        ' Número de SS/DNI, 
        mstrCadenaFinal &= FormateoNoSSDNI(pobjImage) & "|"
        ' Apellidos y Nombre
        mstrCadenaFinal &= FormateoNombreApellidos(pobjImage) & "||"
        ' Fecha de nacimiento
        mstrCadenaFinal &= FormateaFechaNac(pobjImage.VirtualFields.GetFieldValue("DFechaNac").Trim) & "|"
        ' Sexo
        If pobjImage.VirtualFields.GetFieldValue("DSexo") = "1" Then
            pobjImage.VirtualFields.SetFieldValue("DSexo", "M")
        ElseIf pobjImage.VirtualFields.GetFieldValue("DSexo") = "2" Then
            pobjImage.VirtualFields.SetFieldValue("DSexo", "F")
        Else
            pobjImage.VirtualFields.SetFieldValue("DSexo", "U")
        End If
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DSexo") & "||"
        ' Dirección
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DDireccion") & "||"
        ' Teléfono
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DTelefono") & "|"
        ' Doctor
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DDoctor") & "|"
        ' Cargo + Factura
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DFactura") & "||||"
        'mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DFactura") & "^" & pobjImage.VirtualFields.GetFieldValue("DCargo") & "||||"
        ' Diagnóstico
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DCDiagnostico") & "|||"
        ' Número episodio
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DEpisodio") & "|"
        ' Observaciones
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DObservaciones") & "|"

        ' Fecha registro ^ Fecha extracción
        ' PVT, 06/09/2010: Llamamos a la función FormateoFechaRegistroExtraccion que construye la cadena con la fecha de registro y extracción en función de los parámetros especificados en el fichero INI

        mstrCadenaFinal &= FormateoFechaRegistroExtraccion(pobjImage, pdtBatchDate)

        ' Prioridad
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DPrioridad") & "|"
        ' Cama
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DCama") & "||"
        ' Tipo motivo 
        mstrCadenaFinal &= FormateoTipoMotivo(pobjImage) & "|||||"
        ' Servicio
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DServicio") & "|"
        ' Origen/Destino
        mstrCadenaFinal &= FormateoOrigenDestino(pobjImage) & "|"
        ' Grupo/Tipo Fisiológico 
        mstrCadenaFinal &= FormateoGrupoTipoFisiologico(pobjImage) & "|||"

        '14.2.2008 JAC incorporamos 3 campos auxiliares mas, OJO, falta definir cuales son estos campos
        ' DemLibre1
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DNoHist2") & "|"
        ' DemLibre2
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("DNoHist3") & "|"
        ' DemLibre3
        mstrCadenaFinal &= pobjImage.VirtualFields.GetFieldValue("") & vbCrLf

        ' Comentarios de nivel 1 OJO, falta definir cual es este campo!!!!
        If pobjImage.VirtualFields.GetFieldValue("").Trim.Length > 0 Then
            mstrCadenaFinal &= "C|1||" & pobjImage.VirtualFields.GetFieldValue("") & vbCrLf
        End If

    End Sub

    ' ********************************************************************************
    ' FormateoGrupoTipoFisiologico
    ' Desc: Formatea los origenes y destino
    ' NBL: 20/6/2007
    ' ********************************************************************************
    Private Function FormateoGrupoTipoFisiologico(ByRef pobjImage As LibFlexibarNETObjects.Image) As String

        Dim lstrGrupo As String, lstrTipo As String

        lstrGrupo = pobjImage.VirtualFields.GetFieldValue("DGrupo").Trim 'Trim(GetFieldText("DEMGrupo"))
        lstrTipo = pobjImage.VirtualFields.GetFieldValue("DTipoFisiologico").Trim  'Trim(GetFieldText("DEMTipoFisiologico"))

        If lstrGrupo <> "" Or lstrTipo <> "" Then
            Return lstrGrupo & "^" & lstrTipo
        Else
            Return ""
        End If

    End Function

    ' ********************************************************************************
    ' FormateoNoSSDNI
    ' Desc: Formatea los NoSSs y DNIs
    ' NBL: 20/6/2007
    ' ********************************************************************************
    Private Function FormateoNoSSDNI(ByRef pobjImage As LibFlexibarNETObjects.Image) As String

        Dim lstrNoSS As String, lstrDNI As String

        lstrNoSS = pobjImage.VirtualFields.GetFieldValue("DNoSS").Trim
        lstrDNI = pobjImage.VirtualFields.GetFieldValue("DDNI").Trim

        '23.06.2008 JAC rectificado el formateo
        If lstrNoSS = "" And lstrDNI = "" Then
            Return ""
        Else
            If lstrDNI = "" Then
                Return lstrNoSS
            Else
                Return lstrNoSS & "^" & lstrDNI
            End If
        End If

    End Function

    ' ********************************************************************************
    ' FormateoTipoMotivo
    ' Desc: Formatea los tipos y motivos
    ' NBL: 20/6/2007
    ' ********************************************************************************
    Private Function FormateoTipoMotivo(ByRef pobjImage As LibFlexibarNETObjects.Image) As String

        Dim lstrTipo As String, lstrMotivo As String

        lstrTipo = pobjImage.VirtualFields.GetFieldValue("DTipo").Trim
        lstrMotivo = pobjImage.VirtualFields.GetFieldValue("DMotivo").Trim

        '23.06.2008 JAC rectificado el formateo
        If lstrTipo = "" And lstrMotivo = "" Then
            Return ""
        Else
            If lstrMotivo = "" Then
                Return lstrTipo
            Else
                Return lstrTipo & "^" & lstrMotivo
            End If
        End If

    End Function


    ' ********************************************************************************
    ' FormateoOrigenDestino
    ' Desc: Formatea los origenes y destino
    ' NBL: 20/6/2007
    ' ********************************************************************************
    Private Function FormateoOrigenDestino(ByRef pobjImage As LibFlexibarNETObjects.Image) As String

        Dim lstrOrigen As String, lstrDestino As String

        lstrOrigen = pobjImage.VirtualFields.GetFieldValue("DOrigen").Trim
        lstrDestino = pobjImage.VirtualFields.GetFieldValue("DDestino").Trim

        '23.06.2008 JAC rectificado el formateo
        If lstrOrigen = "" And lstrDestino = "" Then
            Return ""
        Else
            If lstrDestino = "" Then
                Return lstrOrigen
            Else
                Return lstrOrigen & "^" & lstrDestino
            End If
        End If

    End Function


    ' ********************************************************************************
    ' FormateoNombreApellidos
    ' Desc: Formatea los nombres y apellidos del paciente
    ' NBL: 20/6/2007
    ' ********************************************************************************
    Private Function FormateoNombreApellidos(ByRef pobjImage As LibFlexibarNETObjects.Image) As String

        Dim lstrNombre As String, lstrApellido1 As String, lstrApellido2 As String

        lstrNombre = pobjImage.VirtualFields.GetFieldValue("DNombre").Trim
        lstrApellido1 = pobjImage.VirtualFields.GetFieldValue("DApellido1").Trim
        lstrApellido2 = pobjImage.VirtualFields.GetFieldValue("DApellido2").Trim

        If lstrNombre <> "" Or lstrApellido1 <> "" Or lstrApellido2 <> "" Then
            Return (lstrApellido1.Trim & " " & lstrApellido2.Trim).Trim & "^^" & lstrNombre.Trim
        Else
            Return ""
        End If

    End Function

    ' ********************************************************************************
    ' FormateoFechasRegistroExtraccion
    ' Desc: Formatea las fechas de registro y extracción dependiendo de los parámetros especificados en el fichero .ini
    ' PVT: 06/09/2010
    ' ********************************************************************************

    Private Function FormateoFechaRegistroExtraccion(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pdtBatchDate As Date) As String

        Dim lstrFechaRegistro As String
        Dim lstrFechaExtraccion As String

        Const C_FORMATO_FECHAHORA As String = "yyyyMMddHHmmss"
        Const C_FORMATO_FECHA As String = "yyyyMMdd"
        Const C_FORMATO_HORA As String = "HHmmss"

        Dim ldtNow As Date = Now

        ' comprobamos si debe aplicarse la configuración standard o la configuración própia de la plantilla.

        If Me.mbolExisteConfiguracionParticular Then ' En el caso de que esté activada la configuración particular de la plantilla.

            ' Comprobamos si debemos o no añadirle la hora a la fecha de Registro.

            If Not Me.mbolPrtAddHoraToFRegistro Then

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de registro

                If Me.mbolPrtBatchDateToFRegistro Then
                    lstrFechaRegistro = Format(pdtBatchDate, C_FORMATO_FECHA)
                Else
                    lstrFechaRegistro = Format(ldtNow, C_FORMATO_FECHA)
                End If

            Else

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de registro

                If Me.mbolPrtBatchDateToFRegistro Then
                    lstrFechaRegistro = Format(pdtBatchDate, C_FORMATO_FECHA) & Format(ldtNow, C_FORMATO_HORA)
                Else
                    lstrFechaRegistro = Format(ldtNow, C_FORMATO_FECHAHORA)
                End If
            End If


            If Not Me.mbolPrtAddHoraToFExtraccion Then

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de extracción

                If Me.mbolPrtBatchDateToFExtraccion Then
                    lstrFechaExtraccion = Format(pdtBatchDate, C_FORMATO_FECHA)
                Else
                    lstrFechaExtraccion = Format(ldtNow, C_FORMATO_FECHA)
                End If

            Else

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de extracción

                If Me.mbolPrtBatchDateToFExtraccion Then
                    lstrFechaExtraccion = Format(pdtBatchDate, C_FORMATO_FECHA) & Format(ldtNow, C_FORMATO_HORA)
                Else
                    lstrFechaExtraccion = Format(ldtNow, C_FORMATO_FECHAHORA)
                End If
            End If

            Return (lstrFechaRegistro & "^" & lstrFechaExtraccion & "|").Trim

        Else ' Configuración standard

            ' Comprobamos si debemos o no añadirle la hora a la fecha de Registro.

            If Not Me.mbolStdAddHoraToFRegistro Then

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de registro

                If Me.mbolStdBatchDateToFRegistro Then
                    lstrFechaRegistro = Format(pdtBatchDate, C_FORMATO_FECHA)
                Else
                    lstrFechaRegistro = Format(ldtNow, C_FORMATO_FECHA)
                End If

            Else

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de registro

                If Me.mbolStdBatchDateToFRegistro Then
                    lstrFechaRegistro = Format(pdtBatchDate, C_FORMATO_FECHA) & Format(ldtNow, C_FORMATO_HORA)
                Else
                    lstrFechaRegistro = Format(ldtNow, C_FORMATO_FECHAHORA)
                End If
            End If


            If Not Me.mbolStdAddHoraToFExtraccion Then

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de extracción

                If Me.mbolStdBatchDateToFExtraccion Then
                    lstrFechaExtraccion = Format(pdtBatchDate, C_FORMATO_FECHA)
                Else
                    lstrFechaExtraccion = Format(ldtNow, C_FORMATO_FECHA)
                End If

            Else

                ' Comprobamos si obtenemos la fecha del lote o bien del sistema a la fecha de extracción

                If Me.mbolStdBatchDateToFExtraccion Then
                    lstrFechaExtraccion = Format(pdtBatchDate, C_FORMATO_FECHA) & Format(ldtNow, C_FORMATO_HORA)
                Else
                    lstrFechaExtraccion = Format(ldtNow, C_FORMATO_FECHAHORA)
                End If
            End If

            Return (lstrFechaRegistro & "^" & lstrFechaExtraccion & "|").Trim

        End If


    End Function


End Class

