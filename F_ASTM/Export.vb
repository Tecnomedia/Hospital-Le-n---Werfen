Imports System.Data.Odbc
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Export

    Private mobjFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch
    Private mobjUtilImage As LibImage.Utils
    Private mstrCarpetaImagenes As String = ""
    Private mstrCarpetaLogImagenes As String = ""
    Private mstrCarpetaLogASTM As String = ""
    Private mstrCarpetaASTM As String = ""
    Private mobjINI As UtilGlobal.clsINI

    Private mintExportLogASTM As Integer = 0
    Private mintExportLogImagenes As Integer = 0

    Private mobjTraductorMarcas As New F_Util.MarkToCode

    ' *******************************************************************************************************
    ' CreaCamposVirtualesExport
    ' Desc: Rutina que crea las variables virtuales que son necesarios para la exportación
    ' NBL 22/4/2009
    ' *******************************************************************************************************
    Private Sub CreaCamposVirtualesExport(ByRef pobjVirtualFields As LibFlexibarNETObjects.colVirtualFields)

        If Not pobjVirtualFields.FieldExists("DNoPet") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoPet", ""))
        If Not pobjVirtualFields.FieldExists("DNoPet2") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoPet2", ""))
        If Not pobjVirtualFields.FieldExists("DNoPet3") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoPet3", ""))
        If Not pobjVirtualFields.FieldExists("DNoHist") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoHist", ""))
        If Not pobjVirtualFields.FieldExists("DNoHist2") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoHist2", ""))
        If Not pobjVirtualFields.FieldExists("DNoHist3") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoHist3", ""))
        If Not pobjVirtualFields.FieldExists("DNoHistFusion") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoHistFusion", ""))
        If Not pobjVirtualFields.FieldExists("DEpisodio") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DEpisodio", ""))
        If Not pobjVirtualFields.FieldExists("DActo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DActo", ""))
        If Not pobjVirtualFields.FieldExists("DNoSS") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNoSS", ""))
        If Not pobjVirtualFields.FieldExists("DDNI") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DDNI", ""))
        If Not pobjVirtualFields.FieldExists("DApellido1") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DApellido1", ""))
        If Not pobjVirtualFields.FieldExists("DApellido2") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DApellido2", ""))
        If Not pobjVirtualFields.FieldExists("DNombre") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNombre", ""))
        If Not pobjVirtualFields.FieldExists("DFechaNac") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFechaNac", ""))
        If Not pobjVirtualFields.FieldExists("DSexo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DSexo", ""))
        If Not pobjVirtualFields.FieldExists("DDireccion") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DDireccion", ""))
        If Not pobjVirtualFields.FieldExists("DTelefono") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DTelefono", ""))
        If Not pobjVirtualFields.FieldExists("DPoblacion") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPoblacion", ""))
        If Not pobjVirtualFields.FieldExists("DCPostal") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DCPostal", ""))
        If Not pobjVirtualFields.FieldExists("DDoctor") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DDoctor", ""))
        If Not pobjVirtualFields.FieldExists("DTDoctor") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DTDoctor", ""))
        If Not pobjVirtualFields.FieldExists("DFactura") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFactura", ""))
        If Not pobjVirtualFields.FieldExists("DCDiagnostico") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DCDiagnostico", ""))
        If Not pobjVirtualFields.FieldExists("DTDiagnostico") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DTDiagnostico", ""))
        If Not pobjVirtualFields.FieldExists("DPrioridad") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPrioridad", ""))
        If Not pobjVirtualFields.FieldExists("DCama") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DCama", ""))
        If Not pobjVirtualFields.FieldExists("DTipo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DTipo", ""))
        If Not pobjVirtualFields.FieldExists("DMotivo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DMotivo", ""))
        If Not pobjVirtualFields.FieldExists("DServicio") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DServicio", ""))
        If Not pobjVirtualFields.FieldExists("DOrigen") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DOrigen", ""))
        If Not pobjVirtualFields.FieldExists("DDestino") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DDestino", ""))
        If Not pobjVirtualFields.FieldExists("DGrupo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DGrupo", ""))
        If Not pobjVirtualFields.FieldExists("DTipoFisiologico") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DTipoFisiologico", ""))
        If Not pobjVirtualFields.FieldExists("DFormID") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFormID", ""))
        If Not pobjVirtualFields.FieldExists("DObservaciones") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DObservaciones", ""))
        If Not pobjVirtualFields.FieldExists("DFHExtraccion") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFHExtraccion", ""))
        If Not pobjVirtualFields.FieldExists("DFHRegistro") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFHRegistro", ""))
        If Not pobjVirtualFields.FieldExists("DPruebas") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPruebas", ""))
        If Not pobjVirtualFields.FieldExists("DPerfiles") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPerfiles", ""))
        If Not pobjVirtualFields.FieldExists("DMuestra") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DMuestra", ""))
        If Not pobjVirtualFields.FieldExists("DResultados") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DResultados", ""))
        If Not pobjVirtualFields.FieldExists("DNTelefono") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DNTelefono", ""))
        If Not pobjVirtualFields.FieldExists("DFaxResultados") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DFaxResultados", ""))
        If Not pobjVirtualFields.FieldExists("DScanStation") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DScanStation", ""))
        If Not pobjVirtualFields.FieldExists("DBatchNo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DBatchNo", ""))
        If Not pobjVirtualFields.FieldExists("DPageNo") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPageNo", ""))
        If Not pobjVirtualFields.FieldExists("DUser") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DUser", ""))

        ' NBL 29/5/2009 Creo un nuevo campo virtual que contendrá el ICU
        'If Not pobjVirtualFields.FieldExists("DICU") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DICU", ""))
        'If Not pobjVirtualFields.FieldExists("DPes") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPes", ""))
        '' NBL 9/6/2009 Creo un nuevo campo para el panel que se ha de abrir en la aplicación de Modulab
        'If Not pobjVirtualFields.FieldExists("DPanel") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("DPanel", ""))

        If Not pobjVirtualFields.FieldExists("ACodigosMarcas") Then pobjVirtualFields.Add(New LibFlexibarNETObjects.VirtualField("ACodigosMarcas", ""))

    End Sub

    ' *******************************************************************************************************
    ' ExportASTM
    ' Desc: Rutina principal de exportación a ASTM
    ' NBL 30/06/2009
    ' *******************************************************************************************************
    Public Function ExportASTM(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                    ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                    ByRef pExportProcessResult As LibFlexibarNETObjects.ExportProcessResult) As Boolean

        mobjFlexibarBatch = pFlexibarBatch

        InicializarExportacion()

        RutinaPrincipalExport(pFlexibarApp, pFlexibarBatch, pExportProcessResult)

    End Function

    ' *******************************************************************************************************
    ' RutinaPrincipalExport
    ' Desc: Inicializamos la configuración para la exportación del ASTM
    ' NBL 30/06/2009
    ' *******************************************************************************************************
    Private Sub RutinaPrincipalExport(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                        ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                        ByRef pExportProcessResult As LibFlexibarNETObjects.ExportProcessResult)

        For lintContador As Integer = 0 To pFlexibarBatch.Images.Count - 1

            Dim lobjImage As LibFlexibarNETObjects.Image = Me.mobjFlexibarBatch.Images(lintContador)

            If lobjImage.RemovePage = False Then
                ExportImageASTM(lobjImage, Me.mobjFlexibarBatch.mobjBatchValues.BatchDate)
            End If

        Next

    End Sub

    ' *******************************************************************************************************
    ' getInfoHistoriaClinica
    ' Desc: Rutina que acondiciona los campos a exportar para exportar los datos
    ' NBL 1/07/2009
    ' *******************************************************************************************************
    Private Sub getInfoHistoriaClinica(ByRef pobjImage As LibFlexibarNETObjects.Image)

        ' 1.- Leyendo de la etiqueta (ICU)
        getInfoHistoriaClinica_Etiqueta(pobjImage)

        ' 2.- Leyendo los bloques de texto y de fecha nacimiento
        'If getInfoHistoriaClinica_Etiqueta(pobjImage) Then Exit Sub

    End Sub

    ' *******************************************************************************************************
    ' getInfoHistoriaClinica_Etiqueta
    ' Desc: Rutina que acondiciona los campos a exportar para exportar los datos
    ' NBL 1/07/2009
    ' *******************************************************************************************************
    Private Sub getInfoHistoriaClinica_Etiqueta(ByRef pobjImage As LibFlexibarNETObjects.Image)

        ' Aquí hay que hacer el tema 
        Dim lstrEtiqueta As String = pobjImage.ARData.ARTextFields.GetFieldValue("A_ETIQUETA")
        Dim lstrFecha As String = ""
        Dim lstrICU As String = ""

        ' Quito los espacios
        lstrEtiqueta = lstrEtiqueta.Replace(" ", "")

        ' Busco la fecha de nacimiento y el ICU en la etiqueta
        Dim lobjMatchFN As Match = Regex.Match(lstrEtiqueta, "\d{2}[/]\d{2}[/]\d{4}")
        Dim lobjMatchICU As Match = Regex.Match(lstrEtiqueta, "ICU:\d{10}")

        If lobjMatchFN.Length > 0 Then
            lstrFecha = lobjMatchFN.Value
        End If
        If lobjMatchICU.Length > 0 Then
            lstrICU = lobjMatchICU.Value.Substring(4)
        End If

        Dim ldtFecha As Date = Nothing
        If lstrFecha.Trim.Length <> 0 And lstrICU.Trim.Length <> 0 Then
            Try
                ldtFecha = CDate(lstrFecha)
            Catch ex As Exception
                Exit Sub
            End Try
            Try
                If Not ConsultaNHC_1(pobjImage, lstrICU, ldtFecha) Then
                    ConsultaNHC_2(pobjImage, lstrICU, ldtFecha)
                End If
            Catch ex2 As Exception
                MsgBox(ex2.ToString)
                Exit Sub
            End Try
        End If

    End Sub

    ' *******************************************************************************************************
    ' ConsultaNHC_1
    ' Desc: Rutina que realiza la consulta de los datos de paciente en el his
    ' NBL 8/10/2009
    ' *******************************************************************************************************
    Private Function ConsultaNHC_1(ByRef pobjImagen As LibFlexibarNETObjects.Image, _
                                            ByVal pstrICU As String, ByVal pdtFechaNac As Date) As Boolean

        Dim lbolDatosEncontrados As Boolean = False

        Dim lstrDSN As String = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "ODBC", "")
        Dim lobjconODBC As New OdbcConnection("DSN=" & lstrDSN)

        'Dim lstrSQL As String = String.Format("SELECT p.numerohc, p.apellid1, p.apellid2, p.nombre, p.numeross1, p.numeross2, p.fechanac, c.servreal, c.fecha " & _
        '                                        "FROM citas AS c, pacientes AS p WHERE (c.numerohc = p.numerohc) " & _
        '                                        "and c.numicu = {0}", pstrICU)

        Dim lstrSQL As String = String.Format("SELECT p.numerohc, p.apellid1, p.apellid2, p.nombre, p.numeross1, p.numeross2, p.fechanac, c.servreal,c.fecha " & _
                                                "FROM citas AS c, pacientes AS p WHERE(c.numerohc = p.numerohc) and c.numicu = '{0}' UNION " & _
                                                "SELECT p.numerohc, p.apellid1, p.apellid2, p.nombre, p.numeross1, p.numeross2, p.fechanac, c.servreal,c.fecha FROM actividad AS c, pacientes AS p " & _
                                                "WHERE(c.numerohc = p.numerohc) and c.numicu = '{0}'", pstrICU)

        lobjconODBC.Open()
        Dim lobjCommand As New OdbcCommand(lstrSQL, lobjconODBC)
        Dim lobjDataReader As OdbcDataReader = lobjCommand.ExecuteReader()
        Dim ldtFechaNacHIS As Date = Nothing

        If lobjDataReader.HasRows Then
            lbolDatosEncontrados = True
            lobjDataReader.Read()
            If Not lobjDataReader.IsDBNull(6) Then
                ldtFechaNacHIS = lobjDataReader.GetDate(6)
            End If
            If ldtFechaNacHIS = pdtFechaNac Then
                ' Se trata del paciente que toca
                If Not lobjDataReader.IsDBNull(0) Then
                    pobjImagen.VirtualFields.SetFieldValue("DNoHist", lobjDataReader.GetInt32(0).ToString())
                End If
                Dim lstrApellidos As String = ""
                If Not lobjDataReader.IsDBNull(1) Then
                    lstrApellidos = lobjDataReader.GetString(1)
                End If
                If Not lobjDataReader.IsDBNull(2) Then
                    lstrApellidos &= " " & lobjDataReader.GetString(2)
                End If
                pobjImagen.VirtualFields.SetFieldValue("DApellido1", lstrApellidos)
                If Not lobjDataReader.IsDBNull(3) Then
                    pobjImagen.VirtualFields.SetFieldValue("DNombre", lobjDataReader.GetString(3))
                End If
                If Not lobjDataReader.IsDBNull(7) Then
                    pobjImagen.VirtualFields.SetFieldValue("DServicio", lobjDataReader.GetString(7))
                End If
                pobjImagen.VirtualFields.SetFieldValue("DFechaNac", Format$(ldtFechaNacHIS, "yyyyMMdd"))
                pobjImagen.VirtualFields.SetFieldValue("DTipo", "CE")
            End If

        End If

        lobjDataReader.Close()
        lobjconODBC.Close()

        Return lbolDatosEncontrados

    End Function

    ' *******************************************************************************************************
    ' ConsultaNHC_2
    ' Desc: Rutina que realiza la consulta de los datos de paciente en el his
    ' NBL 8/10/2009
    ' *******************************************************************************************************
    Private Function ConsultaNHC_2(ByRef pobjImagen As LibFlexibarNETObjects.Image, _
                                            ByVal pstrICU As String, ByVal pdtFechaNac As Date) As Boolean

        Dim lbolDatosEncontrados As Boolean = False

        Dim lstrDSN As String = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "ODBC", "")
        Dim lobjconODBC As New OdbcConnection("DSN=" & lstrDSN)

        Dim lstrSQL As String = String.Format("SELECT  p.numerohc, p.apellid1, p.apellid2, p.nombre, p.numeross1, p.numeross2, p.fechanac, c.servreal, c.fecha_ingreso, c.ncama " & _
                                                                    "FROM camas_paci AS c, pacientes AS p WHERE c.numerohc = p.numerohc and c.numicu = {0}", pstrICU)

        lobjconODBC.Open()
        Dim lobjCommand As New OdbcCommand(lstrSQL, lobjconODBC)
        Dim lobjDataReader As OdbcDataReader = lobjCommand.ExecuteReader()
        Dim ldtFechaNacHIS As Date = Nothing

        If lobjDataReader.HasRows Then
            lbolDatosEncontrados = True
            lobjDataReader.Read()
            If Not lobjDataReader.IsDBNull(6) Then
                ldtFechaNacHIS = lobjDataReader.GetDate(6)
            End If
            If ldtFechaNacHIS = pdtFechaNac Then
                ' Se trata del paciente que toca
                If Not lobjDataReader.IsDBNull(0) Then
                    pobjImagen.VirtualFields.SetFieldValue("DNoHist", lobjDataReader.GetInt32(0).ToString())
                End If
                Dim lstrApellidos As String = ""
                If Not lobjDataReader.IsDBNull(1) Then
                    lstrApellidos = lobjDataReader.GetString(1)
                End If
                If Not lobjDataReader.IsDBNull(2) Then
                    lstrApellidos &= " " & lobjDataReader.GetString(2)
                End If
                pobjImagen.VirtualFields.SetFieldValue("DApellido1", lstrApellidos)
                If Not lobjDataReader.IsDBNull(3) Then
                    pobjImagen.VirtualFields.SetFieldValue("DNombre", lobjDataReader.GetString(3))
                End If
                If Not lobjDataReader.IsDBNull(7) Then
                    pobjImagen.VirtualFields.SetFieldValue("DServicio", lobjDataReader.GetString(7))
                End If
                pobjImagen.VirtualFields.SetFieldValue("DFechaNac", Format$(ldtFechaNacHIS, "yyyyMMdd"))
                pobjImagen.VirtualFields.SetFieldValue("DTipo", "IN")
                If Not lobjDataReader.IsDBNull(9) Then
                    pobjImagen.VirtualFields.SetFieldValue("DCama", lobjDataReader.GetString(9))
                End If
            End If

        End If

        lobjDataReader.Close()
        lobjconODBC.Close()

        Return lbolDatosEncontrados

    End Function


    ' *******************************************************************************************************
    ' getInfoHistoriaClinica_OCR
    ' Desc: Rutina que acondiciona los campos a exportar para exportar los datos
    ' NBL 1/07/2009
    ' *******************************************************************************************************
    Private Function getInfoHistoriaClinica_OCR(ByRef pobjImage As LibFlexibarNETObjects.Image) As Boolean

        Dim lstrNHC_OCR As String = pobjImage.ARData.ARTextFields.GetFieldValue("A_NUMERO_HISTORIA")
        Dim lstrDia As String = pobjImage.ARData.ARTextFields.GetFieldValue("A_DIA_NAC")
        Dim lstrMes As String = pobjImage.ARData.ARTextFields.GetFieldValue("A_MES_NAC")
        Dim lstrAny As String = pobjImage.ARData.ARTextFields.GetFieldValue("A_ANY_NAC")

        If Not IsNumeric(lstrDia) Or Not IsNumeric(lstrMes) Or Not IsNumeric(lstrAny) Then Return False

        Dim lintDia As Integer = CInt(lstrDia), lintMes As Integer = CInt(lstrMes), lintAny As Integer = CInt(lstrAny), lintAnyComplet As Integer = 0

        If lintAny > CInt(Now.Year.ToString.Substring(2, 2)) Then
            lintAnyComplet = lintAny + 1900
        Else
            lintAnyComplet = lintAny + 2000
        End If

        Dim lstrFNac_OCR As String = ""

    End Function

    ' *******************************************************************************************************
    ' AcondicionaVariablesExport
    ' Desc: Rutina que acondiciona los campos a exportar para exportar los datos
    ' NBL 1/07/2009
    ' *******************************************************************************************************
    Private Sub AcondicionaVariablesExport(ByRef pobjImage As LibFlexibarNETObjects.Image)

        ' NBL 2/7/2009  Recupero el número de petición
        If pobjImage.CoreSeparators.Count = 1 Then
            Dim lstrValorSeparador As String = pobjImage.CoreSeparators(0).Value
            If lstrValorSeparador.Length = 8 Then
                pobjImage.VirtualFields.SetFieldValue("DNoPet", lstrValorSeparador)
            ElseIf lstrValorSeparador.Length = 10 Then
                pobjImage.VirtualFields.SetFieldValue("DNoPet", lstrValorSeparador.Substring(1, 8))
            End If
        End If

        If pobjImage.ARData.TemplateName Is Nothing Then Exit Sub
        If pobjImage.ARData.TemplateName.Trim.Length = 9 Then Exit Sub

        ' Recupero la info del médico ----------------------------------------------------------------------------------------------
        Dim lstrMedico As String = mobjTraductorMarcas.ArrayTraductor(pobjImage.ARData.ARCheckmarkFields, "M01", False)
        ' Pongo el valor que corresponda
        If lstrMedico.Length = 5 And lstrMedico <> "00000" Then
            pobjImage.VirtualFields.SetFieldValue("DDoctor", lstrMedico)
        ElseIf lstrMedico.Length < 5 Then
            pobjImage.VirtualFields.SetFieldValue("DDoctor", Microsoft.VisualBasic.Right("00000" & lstrMedico, 5))
        End If


        ' Recupero la información del origen

        ' Recupero la información de la historia clínica
        getInfoHistoriaClinica(pobjImage)

        ' Preparo la variable de códigos de prueba
        pobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", mobjTraductorMarcas.MarkTraductor(pobjImage.ARData.ARCheckmarkFields, pobjImage.ARData.TemplateName, False))

        ' Diagnósticos
        pobjImage.VirtualFields.SetFieldValue("DCDiagnostico", mobjTraductorMarcas.MarkTraductorDiagnostic(pobjImage.ARData.ARCheckmarkFields, pobjImage.ARData.TemplateName, False))


    End Sub

    ' *******************************************************************************************************
    ' ExportImageASTM
    ' Desc: Rutina de exportación de la imagen y del ASTM
    ' NBL 30/06/2009
    ' *******************************************************************************************************
    Private Sub ExportImageASTM(ByRef pobjImage As LibFlexibarNETObjects.Image, ByRef pdtFechaBatch As Date)

        CreaCamposVirtualesExport(pobjImage.VirtualFields)

        ' Tenemos que preparar los datos que vamos a utilizar en el ASTM
        AcondicionaVariablesExport(pobjImage)

        Dim lstrCadenaASTM As String = CrearCadenaASTM(pobjImage.VirtualFields)
        Dim lstrCarpetaFecha As String = pdtFechaBatch.Year & Microsoft.VisualBasic.Right("00" & pdtFechaBatch.Month, 2) & Microsoft.VisualBasic.Right("00" & pdtFechaBatch.Day, 2)

        CopiarImagenExportacion(pobjImage, mstrCarpetaImagenes, lstrCarpetaFecha)
        If mintExportLogImagenes = 1 Then CopiarImagenExportacion(pobjImage, mstrCarpetaLogImagenes, lstrCarpetaFecha)
        If mintExportLogASTM = 1 Then CopiarASTMDebug(pobjImage, mstrCarpetaLogASTM, lstrCarpetaFecha, lstrCadenaASTM)
        CopiarASTM(pobjImage, mstrCarpetaASTM, lstrCadenaASTM)

    End Sub

    ' ****************************************************************************************************
    ' CopiarExtensionPx
    ' Desc: Función que copia a destino con extensión Px
    ' NBL 10/6/2009
    ' ****************************************************************************************************
    Private Sub CopiarExtensionPx(ByVal pstrRutaOrigen As String, ByVal pstrRutaFinal As String)

        Dim lintContador As Integer = 2

        Do
            Dim lstrRutaFinal As String = pstrRutaFinal.Substring(0, pstrRutaFinal.Length - 3) & "p" & lintContador.ToString()
            If Not File.Exists(lstrRutaFinal) Then
                File.Copy(pstrRutaOrigen, lstrRutaFinal)
                Exit Do
            End If
            lintContador += 1
        Loop

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
    ' CrearCadenaASTM
    ' Desc: Rutina que crea la cadena ASTM para cada imagen
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearCadenaASTM(ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields) As String

        Dim lstrCabecera As String = CrearCabeceraASTM(pobjVirtuals)
        Dim lstrRequestDate As String = FechaHoraMensaje()
        Dim lstrPaciente As String = CrearLineaPacienteASTM(pobjVirtuals, lstrRequestDate)
        Dim lstrCodigosMarcas As String = pobjVirtuals.GetFieldValue("ACodigosMarcas")
        Dim lstrCodigosManuscritos As String = pobjVirtuals.GetFieldValue("ACodigosManuscrito")

        Dim lstrCodigos As String = lstrCodigosMarcas & "|" & lstrCodigosManuscritos
        Dim lstrLineasCodigos As String = CrearLineasCodigos(lstrCodigos, pobjVirtuals, lstrRequestDate)
        Dim lstrLineaComentarios As String = CrearLineaComentarios(pobjVirtuals)
        Dim lstrTerminador As String = "L|1|N" & vbCr

        Dim lstrASTMTotal As String = lstrCabecera & lstrPaciente & lstrLineaComentarios & lstrLineasCodigos & lstrTerminador

        Return lstrASTMTotal

    End Function

    ' *******************************************************************************************************
    ' CrearLineasCodigos
    ' Desc: Rutina que crea el listado de códigos de prueba
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Function CrearLineasCodigos(ByVal pstrCodigos As String, ByRef pobjVirtuals As LibFlexibarNETObjects.colVirtualFields, ByVal pstrRequestDate As String) As String

        Dim lstrResultado As String = ""
        Dim lstrCodigos() As String = pstrCodigos.Split("|")
        For lintContador As Integer = 0 To lstrCodigos.Length - 1
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
                lstrResultado &= "|" & "^^^" & lstrCodigos(lintContador).ToString()
                'Campo6
                lstrResultado &= "|" & "R"
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
        lstrLineaPaciente &= "|" '& pobjVirtuals.GetFieldValue("DNoHist3") & "^" & pobjVirtuals.GetFieldValue("DNoSS") & "^" & pobjVirtuals.GetFieldValue("DNoHist2")
        'Campo6
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DApellido1") & "^" & pobjVirtuals.GetFieldValue("DNombre")
        'Campo7
        lstrLineaPaciente &= "|"
        'Campo8
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DFechaNac")
        'Campo9
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DSexo")
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
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DTipo")
        'Campo17 
        lstrLineaPaciente &= "|"
        'Campo18
        lstrLineaPaciente &= "|"
        'Campo19
        lstrLineaPaciente &= "|" & pobjVirtuals.GetFieldValue("DCDiagnostico")
        'Campo20
        lstrLineaPaciente &= "|"
        'Campo21
        lstrLineaPaciente &= "|"
        'Campo22
        lstrLineaPaciente &= "|"
        'Campo23
        lstrLineaPaciente &= "|" '& pobjVirtuals.GetFieldValue("DObservaciones")
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
    ' CopiarImagenExportacion
    ' Desc: Rutina que envia la imagen a la carpeta de exportación
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Private Sub CopiarImagenExportacion(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrCarpetaBase As String, ByVal pstrCarpetaFecha As String)

        If Not My.Computer.FileSystem.DirectoryExists(pstrCarpetaBase) Then
            MsgBox("La carpeta de exportación de imágenes no existe: " & pstrCarpetaBase)
            Throw New Exception("La carpeta de exportación de imágenes no existe: " & pstrCarpetaBase)
        End If

        Dim lstrCarpetaFecha As String = pstrCarpetaBase & "\" & pstrCarpetaFecha
        If Not My.Computer.FileSystem.DirectoryExists(lstrCarpetaFecha) Then My.Computer.FileSystem.CreateDirectory(lstrCarpetaFecha)

        Dim lstrNombreImagenFinal As String = lstrCarpetaFecha & "\" & pobjImage.VirtualFields.GetFieldValue("DNoPet") & ".tif"

        If My.Computer.FileSystem.FileExists(lstrNombreImagenFinal) Then
            ' mobjUtilImage.UnirTiFFs(lstrNombreImagenFinal, pobjImage.Path)
            CopiarExtensionPx(pobjImage.Path, lstrNombreImagenFinal)
        Else
            My.Computer.FileSystem.CopyFile(pobjImage.Path, lstrNombreImagenFinal)
        End If

    End Sub

    ' *******************************************************************************************************
    ' InicializarExportacion
    ' Desc: Inicializamos la configuración para la exportación del ASTM
    ' NBL 30/06/2009
    ' *******************************************************************************************************
    Private Sub InicializarExportacion()

        mobjUtilImage = New LibImage.Utils
        ' Nos conectamos al servidor
        mobjINI = New UtilGlobal.clsINI
        'mstrIP = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "Socket", "IP")
        'mstrPort = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "Socket", "Port")
        mstrCarpetaImagenes = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "RutaCarpetaImagenes")
        mstrCarpetaLogImagenes = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "RutaBackupImagenes")
        mstrCarpetaLogASTM = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "RutaLogASTM")
        mstrCarpetaASTM = mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "RutaASTM")

        mintExportLogASTM = CInt(mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "ActivarLogASTM"))
        mintExportLogImagenes = CInt(mobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_ASTM.ini", "General", "ActivarBackupImagenes"))

    End Sub

End Class
