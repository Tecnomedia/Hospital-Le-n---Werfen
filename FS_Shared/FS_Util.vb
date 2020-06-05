Imports System.Data.Odbc

Public Class FS_Util

    ' *****************************************************************************************
    ' RecuperaDatosPacienteHIS
    ' Desc.: Consulta al HIS del hospital de los datos del paciente
    ' NBL 6/10/2011
    ' *****************************************************************************************
    Public Sub RecuperaDatosPaciente(ByRef pobjImage As LibFlexibarNETObjects.Image, ByVal pstrNHC As String)

        Dim lstrODBC As String = ""
        Dim lobjINI As New UtilGlobal.clsINI

        lstrODBC = lobjINI.IniGet(UtilGlobal.UShared.GetFolderLaboCFG & "\LEON_URGENCIAS.ini", "General", "ODBC")

        If lstrODBC.Trim.Length = 0 Then Exit Sub

        Dim lobjConODBC As New OdbcConnection("DSN=" & lstrODBC)

        Try
            lobjConODBC.Open()
        Catch ex As Exception
            Exit Sub
        End Try

        Dim lstrSQL As String = String.Format("SELECT numerohc, nombre, apellid1, apellid2, fechanac, sexo, dni, letranif, numeross1, numeross2, numeross3 FROM pacientes WHERE numerohc = {0}", _
                                              pstrNHC)

        Dim lobjCommand As New OdbcCommand(lstrSQL, lobjConODBC)
        Dim lobjDataReader As OdbcDataReader = lobjCommand.ExecuteReader()

        If lobjDataReader.HasRows Then

            lobjDataReader.Read()

            ' NHC
            If Not lobjDataReader.IsDBNull(0) Then
                pobjImage.VirtualFields.SetFieldValue("DNoHist", lobjDataReader.GetInt32(0).ToString())
            End If
            ' nombre
            If Not lobjDataReader.IsDBNull(1) Then
                pobjImage.VirtualFields.SetFieldValue("DNombre", lobjDataReader.GetString(1).ToString())
            End If
            ' apellido 1
            If Not lobjDataReader.IsDBNull(2) Then
                pobjImage.VirtualFields.SetFieldValue("DApellido1", lobjDataReader.GetString(2).ToString())
            End If
            ' apellido 2
            If Not lobjDataReader.IsDBNull(3) Then
                pobjImage.VirtualFields.SetFieldValue("DApellido2", lobjDataReader.GetString(3).ToString())
            End If
            ' fechanac
            If Not lobjDataReader.IsDBNull(4) Then
                Dim ldtFechaNac As Date = lobjDataReader.GetDate(4)
                Dim lstrFechaNac As String = Microsoft.VisualBasic.Right("00" & ldtFechaNac.Day, 2) & "/" & Microsoft.VisualBasic.Right("00" & ldtFechaNac.Month, 2) & "/" & ldtFechaNac.Year
                pobjImage.VirtualFields.SetFieldValue("DFechaNac", lstrFechaNac)
            End If
            ' sexo
            If Not lobjDataReader.IsDBNull(5) Then
                pobjImage.VirtualFields.SetFieldValue("DSexo", lobjDataReader.GetInt16(5).ToString())
            End If

            ' dni
            Dim lstrDNI As String = ""
            Dim lstrNIF As String = ""
            If Not lobjDataReader.IsDBNull(6) Then
                lstrDNI = lobjDataReader.GetInt32(6).ToString
            End If
            If Not lobjDataReader.IsDBNull(7) Then
                lstrNIF = lobjDataReader.GetString(7)
            End If
            pobjImage.VirtualFields.SetFieldValue("DDNI", lstrDNI & lstrNIF)
            ' num SS
            Dim lstrNSS1 As String = ""
            Dim lstrNSS2 As String = ""
            Dim lstrNSS3 As String = ""
            If Not lobjDataReader.IsDBNull(8) Then
                lstrNSS1 = lobjDataReader.GetInt16(8).ToString
                lstrNSS2 = lobjDataReader.GetInt32(9).ToString
                lstrNSS3 = lobjDataReader.GetInt16(10).ToString
            End If
            pobjImage.VirtualFields.SetFieldValue("DNoSS", lstrNSS1 & lstrNSS2 & lstrNSS3)


        End If

        lobjDataReader.Close()
        lobjConODBC.Close()

        '        Nombre de la tabla:           pacientes

        '       Nombre de los campos:  0   numerohc    integer     -->  Numero de historia clinica
        '                                           1 nombre        char(20)  -->  Nombre del paciente
        '                                         2 apellid1        char(20)  -->  Apellido primero del paciente
        '                                         3 apellid2        char(20)  -->  Apellido segundo del paciente
        '                                         4 fechanac      date         -->  Fecha de nacimiento del paciente
        '                                         5 sexo             smallint    -->  Codigo del sexo del paciente
        '                                         6 edad                            -->  No existe este campo en la tabla
        '                                         7 dni                integer     -->  Numeros del DNI del paciente
        '                                         8 letranif         char(1)    -->  Letra del DNI del paciente
        '                                         9 numeross1   smallint   -->  Primera parte del numero de S.S.
        '                                         10 numeross2   integer    -->  Segunda parte del numero de S.S.
        '                                         11 numeross3   smallint   -->  Tercera parte del numero de S.S.

    End Sub

    ' *********************************************************************************************
    ' ConsultaDestino
    ' Desc: Hacemos la consulta del destino
    ' NBL 8/5/2009
    ' *********************************************************************************************
    Public Sub ConsultaDestino(ByRef pobjImage As LibFlexibarNETObjects.Image, ByRef pobjOmega As LibOmega.ComClsOmega)

        If pobjImage.VirtualFields.FieldExists("DDestinoAUX") And pobjImage.VirtualFields.GetFieldValue("DDestinoAUX").Trim.Length > 0 Then
            Dim lstrDestinoId As String = pobjImage.VirtualFields.GetFieldValue("DDestinoAUX")
            'pobjImage.VirtualFields.SetFieldValue("DDestino", lstrDestinoId)
            Dim lstrResultado As String = pobjOmega.ConsultaDestino(lstrDestinoId)
            If lstrResultado.Trim.Length > 1 Then
                pobjImage.VirtualFields.SetFieldValue("DDestino", lstrDestinoId)
                Dim lstrDatosDestino() As String = lstrResultado.Split("|")
                If lstrDatosDestino(1).Trim.Length > 0 Then
                    pobjImage.VirtualFields.SetFieldValue("TDestino", lstrDatosDestino(1).Trim)
                End If
            Else
                pobjImage.VirtualFields.SetFieldValue("DDestino", "")
            End If
        End If
    End Sub

End Class
