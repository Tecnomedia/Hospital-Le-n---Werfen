Public Class clsCheckmarkComparer

    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim lobjCheckMarkX As LibFlexibarNETObjects.ARCheckmarkField = CType(x, LibFlexibarNETObjects.ARCheckmarkField)
        Dim lobjCheckMarkY As LibFlexibarNETObjects.ARCheckmarkField = CType(y, LibFlexibarNETObjects.ARCheckmarkField)

        Dim lintNegroX As Integer = lobjCheckMarkX.AmountOfBlack
        Dim lintNegroY As Integer = lobjCheckMarkY.AmountOfBlack

        Return lintNegroX.CompareTo(lintNegroY)

    End Function

End Class
