<Serializable()> _
Public Class clsConfig

    Public Conexion As Integer
    Public TipoLocal As Integer
    Public cnLocal As String
    Public DSNRemota As String
    Public reHC As String ' Expresi�n regular de busqueda de historias cl�nicas
    Public TipoConsulta As Integer ' 0: Empieza por... 1: Incluye
    Public MuestraMicro As Integer '0:Vinculados 1:Desvinculados

End Class
