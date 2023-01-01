Public Class listviewcolumnsorter
    Implements IComparer
    Private columntosort As Integer
    Private orderofsort As SortOrder
    Private objectcompare As CaseInsensitiveComparer

    Public Sub New()
        columntosort = 0
        orderofsort = SortOrder.None
        objectcompare = New CaseInsensitiveComparer()
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim compareresult As Integer
        Dim listviewx As ListViewItem
        Dim listviewy As ListViewItem
        listviewx = CType(x, ListViewItem)
        listviewy = CType(y, ListViewItem)
        compareresult = objectcompare.Compare(listviewx.SubItems(columntosort).Text, listviewy.SubItems(columntosort).Text)
        If (orderofsort = SortOrder.Ascending) Then
            Return compareresult
        ElseIf (orderofsort = SortOrder.Descending) Then
            Return (-compareresult)
        Else Return 0
        End If
    End Function

    Public Property sortcolumn() As Integer
        Set(ByVal value As Integer)
            columntosort = value
        End Set
        Get
            Return columntosort
        End Get
    End Property

    Public Property order() As SortOrder
        Set(ByVal value As SortOrder)
            orderofsort = value
        End Set
        Get
            Return orderofsort
        End Get
    End Property
End Class
