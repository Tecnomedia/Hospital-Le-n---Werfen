Imports System.IO
Imports System.Text.Encoding

Public Class frmProgressExport

    Private mobjFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch
    '    Private mobjSocketClient As Chilkat.Socket
    'Private mintImagenProceso As Integer = 0
    Private mobjUtilImage As LibImage.Utils
    Private mstrCarpetaImagenes As String = ""
    Private mstrCarpetaLogImagenes As String = ""
    Private mstrCarpetaLogASTM As String = ""
    Private mstrCarpetaASTM As String = ""
    Private mobjINI As UtilGlobal.clsINI
    Private mstrIP As String = ""
    Private mstrPort As String = ""
    Private mstrRutaEstadoPaciente As String = ""

    Public Sub New(ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch)

        InitializeComponent()
        mobjFlexibarBatch = pFlexibarBatch

        mobjUtilImage = New LibImage.Utils

        ' Nos conectamos al servidor
        mobjINI = New UtilGlobal.clsINI
        'mstrIP = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "Socket", "IP")
        'mstrPort = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "Socket", "Port")
        mstrCarpetaImagenes = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\MODULAB_ASTM.ini", "General", "RutaCarpetaImagenes")
        mstrCarpetaLogImagenes = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\MODULAB_ASTM.ini", "General", "RutaBackupImagenes")
        mstrCarpetaLogASTM = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\MODULAB_ASTM.ini", "General", "RutaLogASTM")
        mstrCarpetaASTM = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\MODULAB_ASTM.ini", "General", "RutaASTM")
        mstrRutaEstadoPaciente = UtilGlobal.UShared.GetFolderLaboCFG & "\EstadoPaciente.txt"

        ' Inicializamos el cliente de socket
        'mobjSocketClient = New Chilkat.Socket()
        'Dim lbolSuccess As Boolean
        'lbolSuccess = mobjSocketClient.UnlockComponent("SOFICINASocket_BavHPQ6ApEOe")
        'If Not lbolSuccess Then
        '    Me.Timer1.Enabled = False
        '    Throw New Exception("No se ha podido conectar al servidor")
        '    Exit Sub
        'End If
        'Dim ssl As Boolean
        'ssl = False
        'Dim maxWaitMillisec As Long
        'maxWaitMillisec = 20000
        'Dim port As Long
        'port = 4410
        'lbolSuccess = mobjSocketClient.Connect(mstrIP, mstrPort, ssl, maxWaitMillisec)
        'If (lbolSuccess <> True) Then
        '    Me.Timer1.Enabled = False
        '    Throw New Exception(mobjSocketClient.LastErrorText)
        '    Exit Sub
        'End If

        ''  Set maximum timeouts for reading an writing (in millisec)
        'mobjSocketClient.MaxReadIdleMs = 10000
        'mobjSocketClient.MaxSendIdleMs = 10000

        Me.prgEstadoProceso.Minimum = 0
        Me.prgEstadoProceso.Maximum = mobjFlexibarBatch.Images.Count
        Me.prgEstadoProceso.Value = 0

    End Sub

    'Private Sub mobjSocketClient_DatosRecibidos(ByVal datos As String) Handles mobjSocketClient.DatosRecibidos

    '    If datos = "OK" Then
    '        System.Windows.Forms.Application.DoEvents()
    '        EscribeLog("Respuesta de servidor recibida")
    '        EscribeLog("Exportación completada")
    '        EscribeLog("")
    '        Me.Timer1.Enabled = True
    '    End If

    'End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Me.Timer1.Enabled = False

        'If Me.mintImagenProceso > Me.mobjFlexibarBatch.Images.Count - 1 Then
        '    EscribeLog("Fin de exportación")
        '    Me.DialogResult = Windows.Forms.DialogResult.OK
        '    Me.Close()
        '    Exit Sub
        'End If

        RutinaPrincipalExportacionASTM()

    End Sub

    ' **********************************************************************************************
    ' EscribeLog
    ' Desc: Rutina que escribimos a log
    ' 21/5/2009
    ' **********************************************************************************************
    Private Sub EscribeLog(ByVal pstrTexto As String)

        Dim lstrHora As String = Microsoft.VisualBasic.Right("00" & Now.Hour, 2) & ":" & Microsoft.VisualBasic.Right("00" & Now.Month, 2) & ":" & Microsoft.VisualBasic.Right("00" & Now.Day, 2)

        Me.lstLogExportacion.Items.Add(lstrHora & " - " & pstrTexto)
        Me.lstLogExportacion.SelectedIndex = Me.lstLogExportacion.Items.Count - 1

    End Sub

    ' **********************************************************************************************
    ' RutinaPrincipalExportacionASTM
    ' Desc: Rutina principal que llamamos y que realiza la exportación ASTM
    ' 20/5/2009
    ' **********************************************************************************************
    Private Sub RutinaPrincipalExportacionASTM()

        For lintContador As Integer = 0 To Me.mobjFlexibarBatch.Images.Count - 1

            Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(lintContador)

            EscribeLog("Iniciando exportacion imagen: " & lintContador)

            Me.prgEstadoProceso.Value += 1

            If lobjImage.RemovePage = False Then
                ExportImageASTM(lobjImage, Me.mobjFlexibarBatch.mobjBatchValues.BatchDate, Me.mobjFlexibarBatch.ParentFolder)
            End If

        Next

        '      mobjSocketClient.Close(20000)

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    ' *******************************************************************************************************
    ' ExportImageASTM
    ' Desc: Rutina que trata cada imagen del batch para crear a ASTM
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Sub ExportImageASTM(ByRef pobjImage As LibFlexibarNETObjects.Image, ByRef pdtFechaBatch As Date, ByVal pstrParentFolder As String)

        Dim lstrCadenaASTM As String = CrearCadenaASTM(pobjImage.VirtualFields)
        Dim lstrCarpetaFecha As String = pdtFechaBatch.Year & Microsoft.VisualBasic.Right("00" & pdtFechaBatch.Month, 2) & Microsoft.VisualBasic.Right("00" & pdtFechaBatch.Day, 2)

        EscribeLog("Exportando imagen")
        CopiarImagenExportacion(pobjImage, mstrCarpetaImagenes, lstrCarpetaFecha, pstrParentFolder)
        EscribeLog("Exportando imagen backup")
        CopiarImagenExportacion(pobjImage, mstrCarpetaLogImagenes, lstrCarpetaFecha, pstrParentFolder)
        EscribeLog("Exportando archivo ASTM backup")
        CopiarASTMDebug(pobjImage, mstrCarpetaLogASTM, lstrCarpetaFecha, lstrCadenaASTM)
        EscribeLog("Enviando ASTM")
        '    EnviarSocket(lstrCadenaASTM)
        CopiarASTM(pobjImage, mstrCarpetaASTM, lstrCadenaASTM)

    End Sub

    ' *******************************************************************************************************
    ' EnviarSocket
    ' Desc: Enviarmos al servidor de socket
    ' NBL 20/05/2009
    ' *******************************************************************************************************
    Private Sub EnviarSocket(ByVal pstrASTM As String)

        'Dim lstrBytes() As Byte = ASCII.GetBytes(pstrASTM)

        ''  Send the byte data.
        'Dim success As Boolean = False
        'success = mobjSocketClient.SendBytes(lstrBytes)
        'If success = False Then
        '    MsgBox(mobjSocketClient.LastErrorText)
        '    Throw New Exception(mobjSocketClient.LastErrorText)
        '    Exit Sub
        'End If
        'EscribeLog("Esperando respuesta")
        'Dim receivedMsg As String
        'receivedMsg = mobjSocketClient.ReceiveString()
        'If (receivedMsg = vbNullString) Then
        '    MsgBox(mobjSocketClient.LastErrorText)
        '    Throw New Exception(mobjSocketClient.LastErrorText)
        '    Exit Sub
        'End If
        'If (receivedMsg <> "ACK") Then
        '    MsgBox("Terminador no válido")
        '    MsgBox(mobjSocketClient.LastErrorText)
        '    Throw New Exception(mobjSocketClient.LastErrorText)
        '    Exit Sub
        'End If
        'EscribeLog("Respuesta OK. Peticion procesada")

    End Sub

    ' *******************************************************************************************************
    ' CopiarASTMDebug
    ' Desc: Rutina que envia el ASTM a un archivo de texto para tener un log de la aplicación
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Sub CopiarASTMDebug(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrCarpetaBase As String, _
                                                    ByVal pstrCarpetaFecha As String, ByVal pstrASTM As String)

        If Not My.Computer.FileSystem.DirectoryExists(pstrCarpetaBase) Then
            MsgBox("La carpeta de archivos ASTM debug no existe: " & pstrCarpetaBase)
            Throw New Exception("La carpeta de archivos ASTM debug no existe: " & pstrCarpetaBase)
        End If

        Dim lstrCarpetaFecha As String = pstrCarpetaBase & "\" & pstrCarpetaFecha
        If Not My.Computer.FileSystem.DirectoryExists(lstrCarpetaFecha) Then My.Computer.FileSystem.CreateDirectory(lstrCarpetaFecha)

        Dim lstrNombreImagenFinal As String = lstrCarpetaFecha & "\" & pobjImage.VirtualFields.GetFieldValue("DNoPet") & ".pet"

        If My.Computer.FileSystem.FileExists(lstrNombreImagenFinal) Then My.Computer.FileSystem.DeleteFile(lstrNombreImagenFinal)

        'Dim lobjStreamWriter As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(lstrNombreImagenFinal, True)

        Dim lobjStreamWriter As New StreamWriter(lstrNombreImagenFinal, True)

        lobjStreamWriter.Write(pstrASTM)

        lobjStreamWriter.Close()

    End Sub

    ' *******************************************************************************************************
    ' CopiarASTM
    ' Desc: Rutina que envia el ASTM a un archivo de texto para tener un log de la aplicación
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Sub CopiarASTM(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrCarpetaBase As String, _
                                                     ByVal pstrASTM As String)

        If Not My.Computer.FileSystem.DirectoryExists(pstrCarpetaBase) Then
            MsgBox("La carpeta de archivos ASTM no existe: " & pstrCarpetaBase)
            Throw New Exception("La carpeta de archivos ASTM no existe: " & pstrCarpetaBase)
        End If

        Dim lstrNombreImagenFinal As String = pstrCarpetaBase & "\" & pobjImage.VirtualFields.GetFieldValue("DNoPet") & ".txt"

        If My.Computer.FileSystem.FileExists(lstrNombreImagenFinal) Then My.Computer.FileSystem.DeleteFile(lstrNombreImagenFinal)

        'Dim lobjStreamWriter As new StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(lstrNombreImagenFinal, True)
        Dim lobjStreamWriter As New StreamWriter(lstrNombreImagenFinal, True)

        lobjStreamWriter.Write(pstrASTM)

        lobjStreamWriter.Close()

        Dim lbolNombre As Boolean = False
        Dim lintContador As Integer = 0
        Dim lstrReNombradoFinal As String = ""

        Do
            lstrReNombradoFinal = pobjImage.VirtualFields.GetFieldValue("DNoPet") & "-" & lintContador.ToString & ".pet"
            If Not My.Computer.FileSystem.FileExists(pstrCarpetaBase & "\" & lstrReNombradoFinal) Then
                lbolNombre = True
            End If
            lintContador += 1
        Loop While lbolNombre = False

        My.Computer.FileSystem.RenameFile(lstrNombreImagenFinal, lstrReNombradoFinal)

    End Sub

    ' *******************************************************************************************************
    ' CopiarImagenExportacion
    ' Desc: Rutina que envia la imagen a la carpeta de exportación
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Sub CopiarImagenExportacion(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrCarpetaBase As String, ByVal pstrCarpetaFecha As String, ByVal pstrParentFolder As String)

        If Not My.Computer.FileSystem.DirectoryExists(pstrCarpetaBase) Then
            MsgBox("La carpeta de exportación de imágenes no existe: " & pstrCarpetaBase)
            Throw New Exception("La carpeta de exportación de imágenes no existe: " & pstrCarpetaBase)
        End If

        Dim lstrCarpetaFecha As String = pstrCarpetaBase & "\" & pstrCarpetaFecha
        If Not My.Computer.FileSystem.DirectoryExists(lstrCarpetaFecha) Then My.Computer.FileSystem.CreateDirectory(lstrCarpetaFecha)

        Dim lstrNombreImagenFinal As String = lstrCarpetaFecha & "\" & pobjImage.VirtualFields.GetFieldValue("DNoPet") & ".tif"

        If My.Computer.FileSystem.FileExists(lstrNombreImagenFinal) Then
            mobjUtilImage.UnirTiFFs(lstrNombreImagenFinal, pobjImage.Path(pstrParentFolder))
        Else
            My.Computer.FileSystem.CopyFile(pobjImage.Path(pstrParentFolder), lstrNombreImagenFinal)
        End If

    End Sub

    ' *******************************************************************************************************
    ' CrearCadenaASTM
    ' Desc: Rutina que crea la cadena ASTM para cada imagen
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearCadenaASTM(ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields) As String

        Dim lstrCabecera As String = CrearCabeceraASTM(pobjVirtuals)
        Dim lstrRequestDate As String = FechaHoraMensaje()
        Dim lstrPaciente As String = CrearLineaPacienteASTM(pobjVirtuals, lstrRequestDate)        
        'Dim lstrCodigosMarcas As String = pobjVirtuals.GetFieldValue("ACodigosMarcas")
        'Dim lstrCodigosManuscritos As String = pobjVirtuals.GetFieldValue("ACodigosManuscrito")
        'Dim lstrCodigos As String = lstrCodigosMarcas & "|" & lstrCodigosManuscritos
        ' NBL 28/09/2011 -----------------------------------------------------------------------------------------
        Dim lstrCodigos As String = pobjVirtuals.GetFieldValue("ACodigosMarcasHemato")
        ' -------------------------------------------------------------------------------------------------------------
        Dim lstrLineasCodigos As String = CrearLineasCodigos(lstrCodigos, pobjVirtuals, lstrRequestDate)
        Dim lstrLineaComentarios As String = CrearLineaComentarios(pobjVirtuals)
        Dim lstrTerminador As String = "L|1|N" & vbCr

        Dim lstrASTMTotal As String = lstrCabecera & lstrPaciente & lstrLineaComentarios & lstrLineasCodigos & lstrTerminador

        Return lstrASTMTotal

    End Function

    ' *******************************************************************************************************
    ' CrearLineaComentarios
    ' Desc: Rutina que crea la linea de comentarios
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearLineaComentarios(ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields) As String

        Dim lstrLineaComentario As String = ""
        Dim lintContador As Integer = 1

        If pobjVirtuals.FieldExists("DFAguda") Then
            lstrLineaComentario &= "C|" & lintContador.ToString & "|P|FASEA^" & pobjVirtuals.GetFieldValue("DFAguda") & vbCr
            lintContador += 1
        End If

        If pobjVirtuals.FieldExists("DFConvaleciente") Then
            lstrLineaComentario &= "C|" & lintContador.ToString & "|P|FASEC^" & pobjVirtuals.GetFieldValue("DFConvaleciente") & vbCr
            lintContador += 1
        End If

        If pobjVirtuals.GetFieldValue("AOrina24").Trim.Length > 0 Then
            lstrLineaComentario &= "C|" & lintContador.ToString & "|P|DIUS^" & pobjVirtuals.GetFieldValue("AOrina24") & "^" & pobjVirtuals.GetFieldValue("AOrina24") & vbCr
        End If

        Return lstrLineaComentario

    End Function

    ' *******************************************************************************************************
    ' CrearLineasCodigos
    ' Desc: Rutina que crea el listado de códigos de prueba
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearLineasCodigos(ByVal pstrCodigos As String, ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields, ByVal pstrRequestDate As String) As String

        Dim lstrResultado As String = ""
        Dim lstrCodigos() As String = pstrCodigos.Split(",")
        For lintContador As Integer = 1 To lstrCodigos.Length - 1
            If lstrCodigos(lintContador).Trim.Length > 0 Then
                'Campo1
                lstrResultado &= "O"
                'Campo2
                lstrResultado &= "|" & (lintContador + 1).ToString()
                'Campo3
                lstrResultado &= "|" & pobjVirtuals.GetFieldValue("DNoPet")
                'Campo4
                lstrResultado &= "|"
                'Campo5
                Dim lintSombrerito As Integer = lstrCodigos(lintContador).IndexOf("^")
                Dim lstrCodigo As String = lstrCodigos(lintContador).Substring(0, lintSombrerito)
                'lstrResultado &= "|" & "^^^" & lstrCodigos(lintContador).ToString()
                lstrResultado &= "|" & "^^^" & lstrCodigo
                'Campo6
                lstrResultado &= "|" & "S"
                'Campo7
                lstrResultado &= "|" & pstrRequestDate
                'Campo8
                lstrResultado &= "|"
                'Campo9
                lstrResultado &= "|"
                'Campo10
                lstrResultado &= "|"
                'Campo11
                lstrResultado &= "|"
                'Campo12
                lstrResultado &= "|" & "A"
                'Campo13
                lstrResultado &= "|"
                'Campo14
                lstrResultado &= "|"
                'Campo15
                lstrResultado &= "|"
                'Campo16
                lstrResultado &= "|"
                'Campo17
                lstrResultado &= "|" & pobjVirtuals.GetFieldValue("DDoctor")
                'Campo18
                lstrResultado &= "|"
                'Campo19
                lstrResultado &= "|" & pobjVirtuals.GetFieldValue("DOrigen")
                'Campo20
                lstrResultado &= "|"
                'Campo21
                lstrResultado &= "|"
                'Campo22
                lstrResultado &= "|"
                'Campo23
                lstrResultado &= "|"
                'Campo24
                lstrResultado &= "|"
                'Campo25
                lstrResultado &= "|"
                'Campo26
                lstrResultado &= "|" & "F"
                'Campo27
                lstrResultado &= "|"
                'Campo28
                lstrResultado &= "|"
                'Campo29
                lstrResultado &= "|"
                'Campo30
                lstrResultado &= "|" & pobjVirtuals.GetFieldValue("DServicio")
                'Campo31
                lstrResultado &= "|"

                lstrResultado &= vbCr

            End If
        Next

        Return lstrResultado

    End Function

    ' *******************************************************************************************************
    ' CrearCabeceraASTM
    ' Desc: Rutina que crea la cabecera de la cadena ASTM
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearCabeceraASTM(ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields) As String

        Dim lstrCabecera As String = ""

        'Campo1
        lstrCabecera &= "H"
        'Campo2
        lstrCabecera &= "|\^&"
        'Campo3
        lstrCabecera &= "|"
        'Campo4
        lstrCabecera &= "|"
        'Campo5
        lstrCabecera &= "|Flexibar.NET"
        'Campo6
        lstrCabecera &= "|"
        'Campo7
        lstrCabecera &= "|"
        'Campo8
        lstrCabecera &= "|"
        'Campo9
        lstrCabecera &= "|"
        'Campo10
        lstrCabecera &= "|"
        'Campo11
        lstrCabecera &= "|"
        'Campo12
        lstrCabecera &= "|P"
        'Campo13
        lstrCabecera &= "|P"
        'Campo14
        lstrCabecera &= "|" & FechaHoraMensaje()

        ' Final de linea
        lstrCabecera &= vbCr

        Return lstrCabecera

    End Function

    ' *******************************************************************************************************
    ' EstadoPaciente
    ' Desc: Función de devuelve el código de estado de paciente dependiendo del servicio
    ' NBL 26/05/2009
    ' *******************************************************************************************************
    Private Function EstadoPaciente(ByVal pstrServicio As String) As String

        If Not My.Computer.FileSystem.FileExists(mstrRutaEstadoPaciente) Then Return ""

        Dim lobjStreamReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(mstrRutaEstadoPaciente)
        Dim lstrResultado As String = ""

        Do While Not lobjStreamReader.EndOfStream
            Dim lstrLinea() As String = lobjStreamReader.ReadLine.Split("|")
            If lstrLinea.Length = 3 Then
                If lstrLinea(1).ToString = pstrServicio Then
                    lstrResultado = lstrLinea(2)
                    Exit Do
                End If
            End If
        Loop

        lobjStreamReader.Close()

        Return lstrResultado

    End Function

    ' *******************************************************************************************************
    ' CrearLineaPacienteASTM
    ' Desc: Rutina que crea la linea de paciente de la cadena ASTM
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearLineaPacienteASTM(ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields, ByVal pstrRequestDate As String) As String

        Dim lstrLineaPaciente As String = ""

        'Campo1
        lstrLineaPaciente &= "P"
        'Campo2
        lstrLineaPaciente &= "|1"
        'Campo3
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DNoPet")
        'Campo4
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DNoHist").ToString()
        'Campo5
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DNoHist3") & "^" & pobjVirtuals.GetFieldValue("DNoSS") & "^" & pobjVirtuals.GetFieldValue("DNoHist2")
        'Campo6
        lstrLineaPaciente &= "|" & (pobjVirtuals.GetFieldValue("DApellido1").Trim & " " & pobjVirtuals.GetFieldValue("DApellido2").Trim).Trim & "^" & pobjVirtuals.GetFieldValue("DNombre").Trim
        'Campo7
        lstrLineaPaciente &= "|"
        'Campo8
        Dim lstrFechaNac As String = pobjVirtuals.GetFieldValue("DFechaNac")
        If lstrFechaNac.Length = 10 Then
            lstrLineaPaciente &= "|" & lstrFechaNac.Substring(6, 4) & lstrFechaNac.Substring(3, 2) & lstrFechaNac.Substring(0, 2)
        Else
            lstrLineaPaciente &= "|"
        End If

        'Campo9
        If pobjVirtuals.GetFieldValue("DSexo") = "1" Then
            lstrLineaPaciente &= "|M"
        ElseIf pobjVirtuals.GetFieldValue("DSexo") = "2" Then
            lstrLineaPaciente &= "|F"
        Else
            lstrLineaPaciente &= "|"
        End If
        'Campo10
        lstrLineaPaciente &= "|"
        'Campo11
        lstrLineaPaciente &= "|^^"
        'Campo12
        lstrLineaPaciente &= "|"
        'Campo13
        lstrLineaPaciente &= "|"
        'Campo14
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DDoctor")
        'Campo15
        lstrLineaPaciente &= "|"
        'Campo16
        lstrLineaPaciente &= "|"
        'Campo17 
        lstrLineaPaciente &= "|"
        'Campo18
        lstrLineaPaciente &= "|"
        'Campo19
        lstrLineaPaciente &= "|"
        'Campo20
        lstrLineaPaciente &= "|"
        'Campo21
        lstrLineaPaciente &= "|"
        'Campo22
        lstrLineaPaciente &= "|"
        'Campo23
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DObservaciones")
        'Campo24
        lstrLineaPaciente &= "|"
        'Campo25
        lstrLineaPaciente &= "|"
        'Campo26
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DCama")
        'Campo27
        lstrLineaPaciente &= "|"
        'Campo28
        lstrLineaPaciente &= "|"
        'Campo29
        lstrLineaPaciente &= "|"
        'Campo30
        lstrLineaPaciente &= "|"
        'Campo31
        lstrLineaPaciente &= "|"
        'Campo32
        lstrLineaPaciente &= "|"
        'Campo33
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DServicio")
        'Campo34
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DDestino")
        'Campo35
        lstrLineaPaciente &= "|"

        'Fin
        lstrLineaPaciente &= vbCr

        Return lstrLineaPaciente

    End Function

    ' *******************************************************************************************************
    ' FechaHoraMensaje
    ' Desc: Rutina que crea la fecha/hora del mensaje en formato YYYYMMDDHHMMSS
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function FechaHoraMensaje() As String

        Dim lstrResultado As String = ""

        Dim ldtFecha As Date = Me.mobjFlexibarBatch.mobjBatchValues.BatchDate

        lstrResultado &= ldtFecha.Year.ToString() & Microsoft.VisualBasic.Right("00" & ldtFecha.Month.ToString(), 2) & Microsoft.VisualBasic.Right("00" & ldtFecha.Day.ToString(), 2)
        lstrResultado &= Microsoft.VisualBasic.Right("00" & ldtFecha.Hour.ToString(), 2)
        lstrResultado &= Microsoft.VisualBasic.Right("00" & ldtFecha.Minute.ToString(), 2)
        lstrResultado &= Microsoft.VisualBasic.Right("00" & ldtFecha.Second.ToString(), 2)

        Return lstrResultado

    End Function

    ' *******************************************************************************************************
    ' FormateaFechaNacimiento
    ' Desc: Rutina que crea la fecha/hora del mensaje en formato DD/MM/YYYY --> YYYYMMDDHHMMSS
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function FormateaFechaNacimiento(ByVal pstrFecha As String) As String

        Dim lstrResultado As String = ""

        If pstrFecha.Trim.Length = 0 Then Return lstrResultado

        If pstrFecha.Trim.Length = 10 Then
            lstrResultado = pstrFecha.Substring(6, 4) & pstrFecha.Substring(3, 2) & pstrFecha.Substring(0, 2) & "120000"
        End If

        Return lstrResultado

    End Function

End Class