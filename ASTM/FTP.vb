Public Class FTP

    Public Function SendFile(ByVal pstrFile As String, ByVal pstrServer As String, ByVal pstrUser As String, ByVal pstrPassword As String, ByVal pstrFileNameServer As String, ByVal pstrFolders As String) As Boolean
        Dim ftp As New Chilkat.Ftp2()

        Dim success As Boolean

        '  Any string unlocks the component for the 1st 30-days.
        success = ftp.UnlockComponent("SOFICINAFTP_HUbZQcwUpPAk")
        If (success <> True) Then
            Throw New Exception("Error con la licencia de Chillkat")
            Return False
            Exit Function
        End If


        ftp.Hostname = pstrServer
        ftp.Username = pstrUser
        ftp.Password = pstrPassword

        '  The default data transfer mode is "Active" as opposed to "Passive".

        '  Connect and login to the FTP server.
        success = ftp.Connect()
        If (success <> True) Then
            Throw New Exception("Error conectando al servidor ftp especificado")
            Return False
            Exit Function
        End If

        ''  Change to the remote directory where the file will be uploaded.

        success = ftp.ChangeRemoteDir("/")
        If (success <> True) Then
            Throw New Exception("Error accediendo a carpeta: " & "/")
            Return False
            Exit Function
        End If

        Dim strFolders() As String = pstrFolders.Split("\")

        For Each strFolder As String In strFolders
            success = ftp.ChangeRemoteDir(strFolder)
            If (success <> True) Then
                Throw New Exception("Error accediendo a carpeta: " & strFolder)
                Return False
                Exit Function
            End If
        Next


        '  Upload a file.
        Dim localFilename As String
        localFilename = pstrFile
        Dim remoteFilename As String
        remoteFilename = pstrFileNameServer

        success = ftp.PutFile(localFilename, remoteFilename)
        If (success <> True) Then
            Throw New Exception("Error enviando fichero al servidor ftp especificado")
            Return False
            Exit Function
        End If


        ftp.Disconnect()

        Return True

    End Function

    Public Function SendText(ByVal pstrFile As String, ByVal pstrServer As String, ByVal pstrUser As String, ByVal pstrPassword As String, ByVal pstrFileNameServer As String, ByVal pstrFolders As String) As Boolean
        Dim ftp As New Chilkat.Ftp2()

        Dim success As Boolean

        '  Any string unlocks the component for the 1st 30-days.
        success = ftp.UnlockComponent("SOFICINAFTP_HUbZQcwUpPAk")
        If (success <> True) Then
            Throw New Exception("Error con la licencia de Chillkat")
            Return False
            Exit Function
        End If


        ftp.Hostname = pstrServer
        ftp.Username = pstrUser
        ftp.Password = pstrPassword

        '  The default data transfer mode is "Active" as opposed to "Passive".

        '  Connect and login to the FTP server.
        success = ftp.Connect()
        If (success <> True) Then
            Throw New Exception("Error conectando al servidor ftp especificado")
            Return False
            Exit Function
        End If

        ''  Change to the remote directory where the file will be uploaded.

        success = ftp.ChangeRemoteDir("/")
        If (success <> True) Then
            Throw New Exception("Error accediendo a carpeta: " & "/")
            Return False
            Exit Function
        End If

        Dim strFolders() As String = pstrFolders.Split("\")

        For Each strFolder As String In strFolders
            success = ftp.ChangeRemoteDir(strFolder)
            If (success <> True) Then
                Throw New Exception("Error accediendo a carpeta: " & strFolder)
                Return False
                Exit Function
            End If
        Next


        '  Upload a file.
        Dim localFilename As String
        localFilename = pstrFile
        Dim remoteFilename As String
        remoteFilename = pstrFileNameServer

        success = ftp.AppendFileFromTextData(pstrFileNameServer, pstrFile, "ANSI")
        If (success <> True) Then
            Throw New Exception("Error enviando fichero al servidor ftp especificado")
            Return False
            Exit Function
        End If


        ftp.Disconnect()

        Return True

    End Function

    Public Function ExistFile(ByVal pstrFile As String, ByVal pstrServer As String, ByVal pstrUser As String, ByVal pstrPassword As String, ByVal pstrFolders As String) As Boolean
        Dim ftp As New Chilkat.Ftp2()

        Dim success As Boolean

        '  Any string unlocks the component for the 1st 30-days.
        success = ftp.UnlockComponent("SOFICINAFTP_HUbZQcwUpPAk")
        If (success <> True) Then
            Throw New Exception("Error con la licencia de Chillkat")
            Return False
            Exit Function
        End If


        ftp.Hostname = pstrServer
        ftp.Username = pstrUser
        ftp.Password = pstrPassword

        '  The default data transfer mode is "Active" as opposed to "Passive".

        '  Connect and login to the FTP server.
        success = ftp.Connect()
        If (success <> True) Then
            Throw New Exception("Error conectando al servidor ftp especificado")
            Return False
            Exit Function
        End If

        ''  Change to the remote directory where the file will be uploaded.

        success = ftp.ChangeRemoteDir("/")
        If (success <> True) Then
            Throw New Exception("Error accediendo a carpeta: " & "/")
            Return False
            Exit Function
        End If

        Dim strFolders() As String = pstrFolders.Split("\")

        For Each strFolder As String In strFolders
            success = ftp.ChangeRemoteDir(strFolder)
            If (success <> True) Then
                Throw New Exception("Error accediendo a carpeta: " & strFolder)
                Return False
                Exit Function
            End If
        Next


        '  Upload a file.
        Dim localFilename As String
        localFilename = pstrFile

        Dim strText As String = ftp.GetRemoteFileTextC(pstrFile, "ANSI")

        Return strText <> ""


        ftp.Disconnect()

        Return True

    End Function

    Public Function RenameRemoteFile(ByVal pstrFile As String, ByVal pstrNewFile As String, ByVal pstrServer As String, ByVal pstrUser As String, ByVal pstrPassword As String, ByVal pstrFolders As String) As Boolean
        Dim ftp As New Chilkat.Ftp2()

        Dim success As Boolean

        '  Any string unlocks the component for the 1st 30-days.
        success = ftp.UnlockComponent("SOFICINAFTP_HUbZQcwUpPAk")
        If (success <> True) Then
            Throw New Exception("Error con la licencia de Chillkat")
            Return False
            Exit Function
        End If

        ftp.Hostname = pstrServer
        ftp.Username = pstrUser
        ftp.Password = pstrPassword

        '  The default data transfer mode is "Active" as opposed to "Passive".

        '  Connect and login to the FTP server.
        success = ftp.Connect()
        If (success <> True) Then
            Throw New Exception("Error conectando al servidor ftp especificado")
            Return False
            Exit Function
        End If

        ''  Change to the remote directory where the file will be uploaded.

        success = ftp.ChangeRemoteDir("/")
        If (success <> True) Then
            Throw New Exception("Error accediendo a carpeta: " & "/")
            Return False
            Exit Function
        End If

        Dim strFolders() As String = pstrFolders.Split("\")

        For Each strFolder As String In strFolders
            success = ftp.ChangeRemoteDir(strFolder)
            If (success <> True) Then
                Throw New Exception("Error accediendo a carpeta: " & strFolder)
                Return False
                Exit Function
            End If
        Next


        '  Upload a file.
        Dim localFilename As String
        localFilename = pstrFile

        Dim strText As String = ftp.RenameRemoteFile(pstrFile, pstrNewFile)

        Return strText <> ""


        ftp.Disconnect()

        Return True

    End Function



End Class

