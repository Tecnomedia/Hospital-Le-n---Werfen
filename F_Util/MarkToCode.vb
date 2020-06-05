Imports LibFlexibarNETObjects
Imports System.IO
Imports System.Text.RegularExpressions

Public Class MarkToCode

    ' *****************************************************************************************************************
    ' ControlCalidadPlantilla
    ' Desc: Función que hace un control de calidad de las marcas del formulario para controlar si está bien anclado o no
    ' NBL 7/4/2010
    ' *****************************************************************************************************************
    Public Function ControlCalidadPlantilla(ByRef pobjMarks As colARCheckmarkFields, ByVal pstrTemplateName As String) As Boolean

        Dim lobjINI As New UtilGlobal.clsINI

        Dim lstrUmbralCalidad As String = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", pstrTemplateName, "UmbralCalidad")

        ' No hay marcas de control, por lo tanto seguimos palante
        If lstrUmbralCalidad = "" Then Return True

        Dim Resultado As Boolean = True

        ' Si la plantilla es nothing lo que hacemos es darlo por false
        If pstrTemplateName Is Nothing Then Return False

        If lstrUmbralCalidad.Trim.Length <> 0 And IsNumeric(lstrUmbralCalidad) Then

            Dim lintUmbralCalidad As Integer = Val(lstrUmbralCalidad)

            ' Hago un bucle por todas las marcas de calidad
            For lintContador As Integer = 0 To pobjMarks.Count - 1
                If Regex.IsMatch(pobjMarks(lintContador).Name, "^(QN|QB)_\d{2}$") Then
                    Dim lintNegroMarca As Integer = pobjMarks(lintContador).AmountOfBlack
                    If pobjMarks(lintContador).Name.StartsWith("QN") Then
                        If lintNegroMarca > lintUmbralCalidad Then
                            pobjMarks(lintContador).Value = "1"
                        Else
                            pobjMarks(lintContador).Value = "0"
                            Resultado = False
                        End If
                    ElseIf pobjMarks(lintContador).Name.StartsWith("QB") Then
                        If lintNegroMarca <= lintUmbralCalidad Then
                            pobjMarks(lintContador).Value = "0"
                        Else
                            pobjMarks(lintContador).Value = "1"
                            Resultado = False
                        End If
                    End If
                End If
            Next

        Else

            Return False

        End If

        Return Resultado

    End Function

    ' *****************************************************************************************************************
    ' CamposNivelNegro
    ' Desc: Función que calcula el nivel de negro de los campos de lectura de negro para ver si hay información
    ' NBL 26/10/2009
    ' *****************************************************************************************************************
    Public Sub CamposNivelNegro(ByRef pobjImage As LibFlexibarNETObjects.Image)

        If pobjImage.ARData.TemplateName Is Nothing Then Exit Sub
        If pobjImage.ARData.TemplateName = "" Then Exit Sub

        Dim lobjINI As New UtilGlobal.clsINI

        Dim lintNumeroCampos As Integer = CInt(lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", pobjImage.ARData.TemplateName, "Numero", 0))
        If lintNumeroCampos = 0 Then Exit Sub

        For lintContador As Integer = 1 To lintNumeroCampos

            Dim lstrClave As String = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", pobjImage.ARData.TemplateName, "Marca_" & Microsoft.VisualBasic.Right("000" & lintContador, 3), "")
            Dim larClave() As String = lstrClave.Split("^")

            If larClave.Length = 2 Then

                Dim lstrNombre As String = larClave(0)
                Dim lintNivelNegro As Integer = larClave(1)
                Dim lobjCheckMark As LibFlexibarNETObjects.ARCheckmarkField = pobjImage.ARData.ARCheckmarkFields(pobjImage.ARData.ARCheckmarkFields.GetFieldIndex(lstrNombre))

                If Not lobjCheckMark Is Nothing Then
                    If lobjCheckMark.AmountOfBlack > lintNivelNegro Then
                        lobjCheckMark.Value = "1"
                    Else
                        lobjCheckMark.Value = "0"
                    End If
                End If

            End If

        Next

    End Sub

    ' *****************************************************************************************************************
    ' ArrayTraductor
    ' Desc: Función que traduce un array de marcas al valor que esté marcado
    ' NBL 28/5/2009
    ' *****************************************************************************************************************
    Public Function ArrayTraductor(ByRef pobjMarks As colARCheckmarkFields, ByVal pstrPrefijoArray As String, ByRef pbolIsSuspicious As Boolean, ByVal pstrTemplateName As String) As String

        ' En primer lugar hemos de filtrar solo a las marcas con el prefijo de la matriz
        Dim lobjMatriz As New colARCheckmarkFields
        Dim lobjINI As New UtilGlobal.clsINI
        Dim lstrSufijo As String = ""

        FiltraMarcasArray(pobjMarks, lobjMatriz, pstrPrefijoArray, lstrSufijo)

        If lobjMatriz.Count = 0 Then Return ""

        ' Cogemos los límites de filas y columnas
        Dim lstrFilas As String = lobjMatriz(0).Name.Substring(7, 2)
        Dim lstrColumnas As String = lobjMatriz(0).Name.Substring(9, 2)
        Dim lintFilas As Integer = CInt(lstrFilas)
        Dim lintColumnas As Integer = CInt(lstrColumnas)

        ' Hacemos un array por el número de columnas que tiene la matriz
        Dim lstrDatos(lintColumnas) As String

        'Dim lintNivelNegroBueno As Integer = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\F_Util.ini", pstrTemplateName, "NivelNegro", 0)

        For lintCountColumn As Integer = 0 To lintColumnas
            Dim lintValorBueno As Integer = -1
            Dim lintNegroUltimo As Integer = 0
            Dim lintNegroMarca As Integer = 0
            Dim lstrUltimoNombreMarca As String = ""

            For lintCountRow As Integer = 0 To lintFilas
                lintNegroMarca = lobjMatriz(lobjMatriz.GetFieldIndex(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas)).AmountOfBlack
                'If lintNegroMarca > lintNivelNegroBueno And lintNegroMarca > lintNegroUltimo Then
                '    lintNegroUltimo = lintNegroMarca
                '    lintValorBueno = lintCountRow
                '    pobjMarks.SetFieldValue(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas, "1")
                'Else
                '    pobjMarks.SetFieldValue(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas, "0")
                'End If
                If lobjMatriz(lobjMatriz.GetFieldIndex(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas)).Value = "1" Then
                    If lintNegroMarca > lintNegroUltimo Then
                        lintNegroUltimo = lintNegroMarca
                        lintValorBueno = lintCountRow
                        pobjMarks.SetFieldValue(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas & lstrSufijo, "1")
                        If lstrUltimoNombreMarca <> "" Then
                            pobjMarks.SetFieldValue(lstrUltimoNombreMarca & lstrSufijo, "0")
                        End If
                        lstrUltimoNombreMarca = pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas
                    Else
                        pobjMarks.SetFieldValue(pstrPrefijoArray & Microsoft.VisualBasic.Right("00" & lintCountRow, 2) & Microsoft.VisualBasic.Right("00" & lintCountColumn, 2) & lstrFilas & lstrColumnas & lstrSufijo, "0")
                    End If
                End If
            Next
            If lintValorBueno = -1 Then
                lstrDatos(lintCountColumn) = ""
            Else
                lstrDatos(lintCountColumn) = lintValorBueno
            End If
        Next

        Dim lstrResultado As String = ""
        For lintCount As Integer = 0 To lintColumnas
            lstrResultado &= lstrDatos(lintCount).ToString()
        Next

        Return lstrResultado

    End Function

    ' *****************************************************************************************************************
    ' FiltraMarcasArrayRE
    ' Desc: Función que filtra las marcas que pertenecen al grupo siempre que cumplan la regular expresión que se pasa por parámetro
    ' NBL 21/5/2010
    ' *****************************************************************************************************************
    Private Sub FiltraMarcasArrayRE(ByRef pobjMarks As colARCheckmarkFields, ByRef pobjMatriz As colARCheckmarkFields, ByVal pstrRegularExpression As String)

        For lintContador As Integer = 0 To pobjMarks.Count - 1
            If Regex.IsMatch(pobjMarks(lintContador).Name, pstrRegularExpression) Then
                pobjMatriz.Add(New ARCheckmarkField(pobjMarks(lintContador).Name.Substring(0, 5), pobjMarks(lintContador).Value,
                                            pobjMarks(lintContador).Left, pobjMarks(lintContador).Right, pobjMarks(lintContador).Top, pobjMarks(lintContador).Bottom,
                                            pobjMarks(lintContador).AmountOfBlack, pobjMarks(lintContador).BlackThreshold, pobjMarks(lintContador).IsSuspicious,
                                            pobjMarks(lintContador).SuspiciousDistance, True))
            End If
        Next

    End Sub

    ' *****************************************************************************************************************
    ' FiltraMarcasArray
    ' Desc: Función que filtra las marcas que pertenecen al array que se indica por el prefijo
    ' NBL 5/5/2009
    ' *****************************************************************************************************************
    Private Sub FiltraMarcasArray(ByRef pobjMarks As colARCheckmarkFields, ByRef pobjMatriz As colARCheckmarkFields, ByVal pstrPrefijoArray As String, ByRef pstrSufijo As String)

        For lintContador As Integer = 0 To pobjMarks.Count - 1
            'If pobjMarks(lintContador).Name.StartsWith(pstrPrefijoArray) Then
            If Regex.IsMatch(pobjMarks(lintContador).Name, "^" & pstrPrefijoArray & "\d{8}") Then
                'pobjMatriz.Add(pobjMarks(lintContador))
                pstrSufijo = pobjMarks(lintContador).Name.Substring(11)
                pobjMatriz.Add(New ARCheckmarkField(pobjMarks(lintContador).Name.Substring(0, 11), pobjMarks(lintContador).Value,
                                            pobjMarks(lintContador).Left, pobjMarks(lintContador).Right, pobjMarks(lintContador).Top, pobjMarks(lintContador).Bottom,
                                            pobjMarks(lintContador).AmountOfBlack, pobjMarks(lintContador).BlackThreshold, pobjMarks(lintContador).IsSuspicious,
                                            pobjMarks(lintContador).SuspiciousDistance, True))
            End If
        Next

    End Sub

    ' *****************************************************************************************************************
    ' FiltraMarcasArray
    ' Desc: Función que filtra las marcas que pertenecen al array que se indica por el prefijo
    ' NBL 5/5/2009
    ' *****************************************************************************************************************
    Private Sub FiltraMarcasArray(ByRef pobjMarks As colARCheckmarkFields, ByRef pobjMatriz As colARCheckmarkFields, ByVal pstrPrefijoArray As String)

        For lintContador As Integer = 0 To pobjMarks.Count - 1
            'If pobjMarks(lintContador).Name.StartsWith(pstrPrefijoArray) Then
            If Regex.IsMatch(pobjMarks(lintContador).Name, "^" & pstrPrefijoArray & "\d{8}") Then
                'pobjMatriz.Add(pobjMarks(lintContador))

                pobjMatriz.Add(New ARCheckmarkField(pobjMarks(lintContador).Name.Substring(0, 11), pobjMarks(lintContador).Value,
                                            pobjMarks(lintContador).Left, pobjMarks(lintContador).Right, pobjMarks(lintContador).Top, pobjMarks(lintContador).Bottom,
                                            pobjMarks(lintContador).AmountOfBlack, pobjMarks(lintContador).BlackThreshold, pobjMarks(lintContador).IsSuspicious,
                                            pobjMarks(lintContador).SuspiciousDistance, True))
            End If
        Next

    End Sub

    ' *******************************************************************************************************
    ' FiltraMarcasHematologia
    ' Desc.: Función que calcula los códigos de prueba de laboratorio de hematología de izasa
    ' NBL 26/09/2011
    ' *******************************************************************************************************
    Private Sub FiltraMarcasHematologia(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrTemplateName As String, _
                                        ByRef pobjPruebasMarked As ArrayList, ByRef pobjMuestrasMarked As ArrayList, ByRef pobjLocaliMarked As ArrayList, _
                                        ByRef pobjAbrvToDesc As Dictionary(Of String, String))

        FilterMarkedHemato(pobjImage.ARData.ARCheckmarkFields, pobjPruebasMarked, pobjMuestrasMarked, pobjLocaliMarked, pobjAbrvToDesc)

    End Sub

    ' ********************************************************************************************************
    ' CalculaPruebasHemato
    ' Desc.: Hacemos el cálculo de pruebas de hematología
    ' NBL 26/9/2011 
    ' ********************************************************************************************************
    Public Sub CalculaPruebasHemato(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lobjPruebasMarked As New ArrayList, lobjMuestrasMarked As New ArrayList, lobjLocaliMarked As New ArrayList
        Dim lobjAbrvToDesc As New Dictionary(Of String, String)

        FiltraMarcasHematologia(pobjImage, pobjImage.ARData.TemplateName, lobjPruebasMarked, lobjMuestrasMarked, lobjLocaliMarked, lobjAbrvToDesc)

        Dim lstrRutaArchivo As String = UtilGlobal.UShared.GetFolderLaboCFG & "\" & "MODULAB_" & pobjImage.ARData.TemplateName.ToString() & ".txt"
        Dim lstrMarcas As String = "", lstrDesc As String = ""
        ProcesaMarcasHematologia(lstrRutaArchivo, lobjPruebasMarked, lstrMarcas, lstrDesc)

        pobjImage.VirtualFields.SetFieldValue("ACodigosMarcasHemato", lstrMarcas)
        pobjImage.VirtualFields.SetFieldValue("ACodigosMarcasHematoDesc", lstrDesc)

    End Sub

    ' ********************************************************************************************************
    ' CalculaPruebasBioquimica
    ' Desc.: Hacemos el cálculo de pruebas de bioquímica
    ' NBL 10/11/2011 
    ' ********************************************************************************************************
    Public Sub CalculaPruebasBioquimica(ByRef pobjImage As LibFlexibarNETObjects.Image)

        Dim lobjPruebasMarked As New ArrayList, lobjMuestrasMarked As New ArrayList, lobjLocaliMarked As New ArrayList
        Dim lobjAbrvToDesc As New Dictionary(Of String, String)

        FiltraMarcasHematologia(pobjImage, pobjImage.ARData.TemplateName, lobjPruebasMarked, lobjMuestrasMarked, lobjLocaliMarked, lobjAbrvToDesc)

        Dim lstrRutaArchivo As String = UtilGlobal.UShared.GetFolderLaboCFG & "\" & "MODULAB_B_" & pobjImage.ARData.TemplateName.ToString() & ".txt"
        Dim lstrMarcas As String = "", lstrDesc As String = ""
        ProcesaMarcasHematologia(lstrRutaArchivo, lobjPruebasMarked, lstrMarcas, lstrDesc)

        pobjImage.VirtualFields.SetFieldValue("ACodigosMarcas", lstrMarcas)
        pobjImage.VirtualFields.SetFieldValue("ACodigosMarcasDesc", lstrDesc)

    End Sub

    ' ********************************************************************************************************
    ' ProcesaMarcasHematologia
    ' Desc.: Función que procesa las marcas a códigos, pasamos por parámetro el archivo de texto para traducir las marcas
    ' NBL 26/9/2011
    ' ********************************************************************************************************
    Private Sub ProcesaMarcasHematologia(ByVal pstrRutaArchivoTraduc As String, ByRef pobjMarks As ArrayList, ByRef pstrMarcas As String, ByRef pstrDesc As String)

        If pobjMarks.Count = 0 Then Exit Sub

        If My.Computer.FileSystem.FileExists(pstrRutaArchivoTraduc) Then

            Dim lobjStreamReader As New StreamReader(pstrRutaArchivoTraduc)
            Dim lstrSeparador As String = lobjStreamReader.ReadLine.Trim

            Do While Not lobjStreamReader.EndOfStream
                Dim lstrLinea As String = lobjStreamReader.ReadLine
                Dim lstrLineaArray() As String = lstrLinea.Split(lstrSeparador)
                If lstrLineaArray.Length = 3 Then
                    If pobjMarks.Contains(lstrLineaArray(0)) Then
                        pstrMarcas &= lstrLineaArray(2)
                        If pstrDesc.Trim.Length = 0 Then
                            pstrDesc = "^" & lstrLineaArray(1).Substring(7) & "#^"
                        Else
                            pstrDesc &= "|^" & lstrLineaArray(1).Substring(7) & "#^"
                        End If
                    End If
                End If
            Loop

            lobjStreamReader.Close()

        End If

    End Sub

    ' ********************************************************************************************************
    ' FilterMarked
    ' Desc.: Función que crea tres colecciones separadas para pruebas, muestras y localizaciones
    ' NBL 26/9/2011
    ' ********************************************************************************************************
    Private Sub FilterMarkedHemato(ByVal pobjMarks As colARCheckmarkFields, ByRef pobjPruebasMarked As ArrayList, _
                             ByRef pobjMuestrasMarked As ArrayList, ByRef pobjLocalizaMarked As ArrayList, _
                             ByRef pobjAbrvToDesc As Dictionary(Of String, String))

        For lintContador As Integer = 0 To pobjMarks.Count - 1

            ' NBL 26/9/2011 Solo tengo en cuenta las pruebas, ya que en hematología no hay ni muestras ni localizaciones
            If Regex.IsMatch(pobjMarks(lintContador).Name, "^P0\d{3}P_") And pobjMarks(lintContador).Value = "1" Then
                pobjPruebasMarked.Add(pobjMarks(lintContador).Name.Substring(0, 6))
                pobjAbrvToDesc.Add(pobjMarks(lintContador).Name.Substring(0, 6), pobjMarks(lintContador).Name)
            End If

        Next

    End Sub

    ' *****************************************************************************************************************
    ' FilterMarked
    ' Desc: Función crea tres colecciones separadas para pruebas, muestras y localizaciones que están marcadas
    ' NBL 5/5/2009
    ' *****************************************************************************************************************
    Private Sub FilterMarked(ByVal pobjMarks As colARCheckmarkFields, ByRef pobjPruebasMarked As ArrayList, _
                                                ByRef pobjPruebasMicroMarked As ArrayList, _
                                                ByRef pobjMuestrasMarked As ArrayList, ByRef pobjLocalizaMarked As ArrayList, _
                                                ByRef pobjAbrvToDesc As Dictionary(Of String, String))

        For lintContador As Integer = 0 To pobjMarks.Count - 1
            If Regex.IsMatch(pobjMarks(lintContador).Name, "^P1\d{3}P") And pobjMarks(lintContador).Value = "1" Then
                pobjPruebasMarked.Add(pobjMarks(lintContador).Name.Substring(0, 6))
                pobjAbrvToDesc.Add(pobjMarks(lintContador).Name.Substring(0, 6), pobjMarks(lintContador).Name)
            ElseIf Regex.IsMatch(pobjMarks(lintContador).Name, "^M1\d{3}P") And pobjMarks(lintContador).Value = "1" Then
                pobjPruebasMicroMarked.Add(pobjMarks(lintContador).Name.Substring(0, 6))
                pobjAbrvToDesc.Add(pobjMarks(lintContador).Name.Substring(0, 6), pobjMarks(lintContador).Name)
            ElseIf Regex.IsMatch(pobjMarks(lintContador).Name, "^M1\d{3}M") And pobjMarks(lintContador).Value = "1" Then
                pobjMuestrasMarked.Add(pobjMarks(lintContador).Name.Substring(0, 6))
                pobjAbrvToDesc.Add(pobjMarks(lintContador).Name.Substring(0, 6), pobjMarks(lintContador).Name)
            ElseIf Regex.IsMatch(pobjMarks(lintContador).Name, "^P1\d{3}L") And pobjMarks(lintContador).Value = "1" Then
                pobjLocalizaMarked.Add(pobjMarks(lintContador).Name.Substring(0, 6))
                pobjAbrvToDesc.Add(pobjMarks(lintContador).Name.Substring(0, 6), pobjMarks(lintContador).Name)
            End If
        Next

    End Sub

    ' *****************************************************************************************************************
    ' MarkTraductor
    ' Desc: Función que traduce las marcas a una colección de virtuales con los códigos asociados
    ' pintTipoMarca: 
    ' NBL 5/5/2009
    ' *****************************************************************************************************************
    Public Function MarkTraductor(ByRef pobjMarks As colARCheckmarkFields, ByVal pstrTemplateName As String, _
                                                ByVal pbolVerificador As Boolean, ByVal pintTipoMarca As Integer, Optional ByRef pstrObs As String = "") As String

        Dim lobjPruebasMarked As New ArrayList, lobjPruebasMicroMarked As New ArrayList, lobjMuestrasMarked As New ArrayList, lobjLocaliMarked As New ArrayList
        Dim lobjAbrvToDesc As New Dictionary(Of String, String)

        ' NBL 19/01/2010 Esto lo pasamos antes para que pueda influir a las matrices de marcas que hay
        'If Not pbolVerificador Then
        '    ModificaValorNivelNegro(pobjMarks, pstrTemplateName)
        'End If

        If pintTipoMarca = 1 Then
            FilterMarked(pobjMarks, lobjPruebasMarked, lobjPruebasMicroMarked, lobjMuestrasMarked, lobjLocaliMarked, lobjAbrvToDesc)
        Else
            ' Aquí solo se utiliza la colección de lobjPruebasMarked
            ' FilterMarkedDiagnostic(pobjMarks, lobjPruebasMarked, lobjMuestrasMarked, lobjLocaliMarked, lobjAbrvToDesc)
        End If

        Select Case pstrTemplateName

            Case "LE_URGENCIAS"

                Dim lstrPruebasBio As String = ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\OMEGA_B_LE_URGENCIAS.txt", lobjPruebasMarked)
                Dim lstrPruebasMicro As String = ProcesaMarcasMicro(UtilGlobal.UShared.GetFolderLaboCFG & "\OMEGA_M_LE_URGENCIAS.txt", lobjMuestrasMarked, lobjPruebasMicroMarked)

                Return lstrPruebasBio & lstrPruebasMicro

                'Case "MIC_SERO_07"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_SERO_07_P.txt", lobjPruebasMarked)
                'Case "MIC_SERO_03"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_SERO_03_P.txt", lobjPruebasMarked)
                'Case "URGENCIAS_07"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\URGENCIAS_07_P.txt", lobjPruebasMarked)
                'Case "PRIMARIA_07"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\PRIMARIA_07_P.txt", lobjPruebasMarked)
                'Case "MIC_BACTERIO_07"
                '    Return ProcesaMarcasMicro("2007", _
                '                                                lobjPruebasMarked, lobjMuestrasMarked)
                '    'Return ProcesaMarcasMicro(UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_BACTERIO_07_P.txt", _
                '    '                                        UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_BACTERIO_07_M.txt", _
                '    '                                        lobjPruebasMarked, lobjMuestrasMarked)
                'Case "MIC_BACTERIO_03"
                '    Return ProcesaMarcasMicro("2003", _
                '                                lobjPruebasMarked, lobjMuestrasMarked)
                '    'Return ProcesaMarcasMicro(UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_BACTERIO_03_P.txt", _
                '    '                                            UtilGlobal.UShared.GetFolderLaboCFG & "\MIC_BACTERIO_03_M.txt", _
                '    '                                            lobjPruebasMarked, lobjMuestrasMarked)
                'Case "BIO_HEMATO_08"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\BIO_HEMATO_08_P.txt", lobjPruebasMarked)
                'Case "ESPECIALIZADA_03"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\ESPECIALIZADA_03_P.txt", lobjPruebasMarked)
                'Case "INMUNOLOGIA_03"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\INMUNOLOGIA_03_P.txt", lobjPruebasMarked)
                'Case "INMUNOLOGIA_07"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\INMUNOLOGIA_07_P.txt", lobjPruebasMarked)
                'Case "URGENCIAS_03"
                '    Return ProcesaMarcas(UtilGlobal.UShared.GetFolderLaboCFG & "\URGENCIAS_03_P.txt", lobjPruebasMarked)

        End Select

        Return ""

    End Function

    ' ********************************************************************************************************
    ' ProcesaMarcasMicro
    ' Desc.: Función que procesa las marcas a códigos, pasamos por parámetro el archivo de texto para traducir las marcas
    ' NBL 27/9/2011
    ' ********************************************************************************************************
    Private Function ProcesaMarcasMicro(ByVal pstrRutaArchivoTraduc As String, ByRef pobjMuestrasMicro As ArrayList, ByRef pobjPruebasMicro As ArrayList) As String

        Dim lstrResultado As String = ""

        ' Si no hay almenos una pareja muestra-prueba
        If pobjMuestrasMicro.Count = 0 Or pobjPruebasMicro.Count = 0 Then Return ""
        If pobjMuestrasMicro.Count > 1 Then Return ""

        Dim lstrCodigoMuestra As String = ""

        If My.Computer.FileSystem.FileExists(pstrRutaArchivoTraduc) Then

            Dim lobjStreamReader As New StreamReader(pstrRutaArchivoTraduc)
            Dim lstrSeparador As String = lobjStreamReader.ReadLine.Trim

            Do While Not lobjStreamReader.EndOfStream
                Dim lstrLinea As String = lobjStreamReader.ReadLine.Trim
                Dim lstrLineaArray() As String = lstrLinea.Split(lstrSeparador)
                If lstrLineaArray.Length = 3 Then
                    If pobjMuestrasMicro.Contains(lstrLineaArray(0)) Then
                        lstrCodigoMuestra = lstrLineaArray(2)
                        Exit Do
                    End If
                End If
            Loop

            lobjStreamReader.Close()

            If lstrCodigoMuestra.Trim.Length = 0 Then Return ""

            If My.Computer.FileSystem.FileExists(pstrRutaArchivoTraduc) Then
                Dim lobjStreamReader_P As New StreamReader(pstrRutaArchivoTraduc)
                Dim lstrSeparador_P As String = lobjStreamReader_P.ReadLine.Trim

                Do While Not lobjStreamReader_P.EndOfStream
                    Dim lstrLinea As String = lobjStreamReader_P.ReadLine
                    Dim lstrLineaArray() As String = lstrLinea.Split(lstrSeparador_P)
                    If lstrLineaArray.Length = 3 Then
                        If pobjPruebasMicro.Contains(lstrLineaArray(0)) Then
                            lstrResultado &= "," & lstrLineaArray(2) & "^M|" & lstrCodigoMuestra & "|"
                        End If
                    End If
                Loop
                lobjStreamReader_P.Close()
            End If

        End If

        Return lstrResultado

    End Function

    ' ********************************************************************************************************
    ' ProcesaMarcas
    ' Desc.: Función que procesa las marcas a códigos, pasamos por parámetro el archivo de texto para traducir las marcas
    ' NBL 26/6/2009
    ' ********************************************************************************************************
    Private Function ProcesaMarcas(ByVal pstrRutaArchivoTraduc As String, ByRef pobjMarks As ArrayList) As String

        Dim lstrResultado As String = ""

        If pobjMarks.Count = 0 Then Return ""

        If My.Computer.FileSystem.FileExists(pstrRutaArchivoTraduc) Then

            Dim lobjStreamReader As New StreamReader(pstrRutaArchivoTraduc)

            Dim lstrSeparador As String = lobjStreamReader.ReadLine.Trim

            Do While Not lobjStreamReader.EndOfStream
                Dim lstrLinea As String = lobjStreamReader.ReadLine
                Dim lstrLineaArray() As String = lstrLinea.Split(lstrSeparador)
                If lstrLineaArray.Length = 3 Then
                    If pobjMarks.Contains(lstrLineaArray(0)) Then
                        lstrResultado &= lstrLineaArray(2)
                    End If
                End If
            Loop

            lobjStreamReader.Close()

        End If

        'If lstrResultado.Length > 0 Then
        '    Return lstrResultado.Substring(1)
        'Else
        '    Return lstrResultado
        'End If

        Return lstrResultado

    End Function

End Class
