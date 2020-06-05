Imports System.Collections
Imports System.Windows.Forms

Friend Class ListViewColumnSorter

    Implements System.Collections.IComparer

    Private ColumnToSort As Integer
    Private OrderOfSort As SortOrder
    Private ObjectCompare As CaseInsensitiveComparer
    Private _OrderDate As Boolean ' Variable que nos indicará que ordenamos por un campo que es fecha

    Public Sub New()
        ' Initialize the column to '0'.
        ColumnToSort = 0

        ' Initialize the sort order to 'none'.
        OrderOfSort = SortOrder.None

        ' Initialize the CaseInsensitiveComparer object.
        ObjectCompare = New CaseInsensitiveComparer()

    End Sub

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements IComparer.Compare

        Dim compareResult As Integer
        Dim listviewX As ListViewItem
        Dim listviewY As ListViewItem

        ' Cast the objects to be compared to ListViewItem objects.
        listviewX = CType(x, ListViewItem)
        listviewY = CType(y, ListViewItem)

        If Not OrderDate Then

            ' Compare the two items.
            compareResult = ObjectCompare.Compare(listviewX.SubItems(ColumnToSort).Text, listviewY.SubItems(ColumnToSort).Text)

            ' Calculate the correct return value based on the object 
            ' comparison.
            If (OrderOfSort = SortOrder.Ascending) Then
                ' Ascending sort is selected, return typical result of 
                ' compare operation.
                Return compareResult
            ElseIf (OrderOfSort = SortOrder.Descending) Then
                ' Descending sort is selected, return negative result of 
                ' compare operation.
                Return (-compareResult)
            Else
                ' Return '0' to indicate that they are equal.
                Return 0
            End If

        Else

            Return CompararFechas(listviewX.SubItems(ColumnToSort).Text, listviewY.SubItems(ColumnToSort).Text)

        End If

    End Function

    Private Function CompararFechas(ByVal pstrFecha1 As String, ByVal pstrFecha2 As String) As Integer

        If IsDate(pstrFecha1) And IsDate(pstrFecha2) Then
            Dim ldtFecha1 As Date = Date.Parse(pstrFecha1)
            Dim ldtFecha2 As Date = Date.Parse(pstrFecha2)
            Return Date.Compare(ldtFecha1, ldtFecha2)
        End If

        If IsDate(pstrFecha1) And Not IsDate(pstrFecha2) Then
            Return 1
        End If

        If Not IsDate(pstrFecha1) And IsDate(pstrFecha2) Then
            Return -1
        End If

        If Not IsDate(pstrFecha1) And Not IsDate(pstrFecha2) Then
            Return 0
        End If

    End Function

    Public Property SortColumn() As Integer
        Set(ByVal Value As Integer)
            ColumnToSort = Value
        End Set

        Get
            Return ColumnToSort
        End Get
    End Property

    Public Property Order() As SortOrder
        Set(ByVal Value As SortOrder)
            OrderOfSort = Value
        End Set

        Get
            Return OrderOfSort
        End Get
    End Property

    Public Property OrderDate() As Boolean
        Get
            Return _OrderDate
        End Get
        Set(ByVal value As Boolean)
            _OrderDate = value
        End Set
    End Property
End Class


