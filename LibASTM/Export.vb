Imports System.IO

Public Class Export

    ' *******************************************************************************************************
    ' ExportASTM
    ' Desc: Rutina principal de exportación a ASTM
    ' NBL 15/05/2009
    ' *******************************************************************************************************
    Public Function ExportASTM(ByRef pFlexibarApp As LibFlexibarNETObjects.FlexibarApp, _
                                                    ByRef pFlexibarBatch As LibFlexibarNETObjects.FlexibarBatch, _
                                                    ByRef pExportProcessResult As LibFlexibarNETObjects.ExportProcessResult) As Boolean

        ' Hemos de hacer un bucle por cada una de las imágenes
        'If pFlexibarBatch.Images.Count > 0 Then
        '    For lintContador As Integer = 0 To pFlexibarBatch.Images.Count - 1
        '        Try
        '            If pFlexibarBatch.Images(lintContador).RemovePage = False Then
        '                ExportImageASTM(pFlexibarBatch.Images(lintContador), pFlexibarBatch.mobjBatchValues.BatchDate)
        '            End If
        '        Catch ex As Exception
        '            pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
        '            pExportProcessResult.ErrorDescription = ex.Source & " - " & ex.Message
        '        End Try
        '    Next
        'End If

        Try
            Dim lobjfrmExport As New frmProgressExport(pFlexibarBatch)
            If lobjfrmExport.ShowDialog() <> Windows.Forms.DialogResult.OK Then
                pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
                pExportProcessResult.ErrorDescription = "Error en la exportación de ASTM"
            End If
        Catch ex As Exception
            pExportProcessResult.ExportStatus = LibFlexibarNETObjects.enExportStatus.ExportError
            pExportProcessResult.ErrorDescription = ex.Source & " - " & ex.Message
        End Try

    End Function



End Class
