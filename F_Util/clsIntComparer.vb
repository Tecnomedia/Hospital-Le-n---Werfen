Public Class clsIntComparer

    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        'Both Numbers
        If IsNumeric(x) And IsNumeric(y) Then
            Return CInt(x).CompareTo(y)
        ElseIf IsNumeric(x) And Not IsNumeric(y) Then
            'X is number Y is Alpha
            Return -1
        ElseIf Not IsNumeric(x) And IsNumeric(y) Then
            'X is Alpha Y is Number
            Return 1
        Else
            'Both Alpha
            Return String.Compare(x, y)
        End If
    End Function

End Class
