Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Module Module1
    Public lvwcolumnsorter As listviewcolumnsorter
    Dim databasefilename As String = Application.StartupPath & "\db.mdb"
    Public connector As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & databasefilename)
    Public request As String
    Public myXL As Excel.Application, myWB As Excel.Workbook, myWS As Excel.Worksheet
    Public id As String
    Public номер As String
    Public IP As String
    Public DNS As String
    Public Марка As String
    Public SN As String
    Public Устройство As String
    Public Наличие As String
    Public Фамилия As String
    Public Имя As String
    Public Отчество As String
    Public Получатель As String
    Public ПК As String
    Public ФИО As String
    Public Комп As String
    Public Ошибка As String

    Public Sub ReadDataFromTable(ByVal request As String, ByVal listview As Object, ByVal sum As Integer)
        Try
            Dim datareader As OleDbDataReader
            Dim command As New OleDbCommand(request, connector)
            Dim kartinka As Integer = 3
            connector.Open()
            datareader = command.ExecuteReader
            While datareader.Read() = True
                Select Case sum
                    Case 4
                        With listview
                            .Items.Add(datareader.GetValue(0))
                            .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(1))
                            .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(2))
                            .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(3))
                        End With
                End Select
            End While
            datareader.Close()
            connector.Close()
        Catch ex As Exception
            connector.Close()
            MsgBox(ex.Message)
        End Try
        listview.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In listview.Columns
            columnheader.width = -3
        Next
    End Sub

    Public Sub ChangeDataInTable(ByVal request As String)
        Try
            Dim command As New OleDbCommand(request, connector)
            connector.Open()
            command.ExecuteNonQuery()
            connector.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Sub ExportWord(ByVal listview As Object)
        Dim Дата As String = Format(Now, "d MMMM yyyy")
        Dim W = New Word.Application
        W.Visible = True
        W.Documents.Add()
        W.Selection.TypeText("Архив выдачи техники на: " & Дата & Chr(13) & Chr(10))
        With listview
            For i As Short = 0 To .Items.Count - 1
                W.Selection.TypeText(.Items(i).SubItems.Item(0).Text & " " & .Items(i).SubItems.Item(1).Text & " " & .Items(i).SubItems.Item(2).Text & Chr(13) & Chr(10) & " " & .Items(i).SubItems.Item(3).Text & Chr(13) & Chr(10))
            Next i
        End With
    End Sub

    Sub ExportExcel(ByVal listview As Object, ByVal sum As Integer)
        Dim i As Integer
        Dim Y As Integer
        Dim z As Integer
        myXL = New Excel.Application
        myWB = myXL.Workbooks.Add
        myWS = myWB.Worksheets(1)
        z = 2
        myXL.Visible = True
        With listview
            Select Case sum
                Case 2
                    For i = 1 To .Items.Count
                        For Y = 1 To 2
                            myWS.Cells(1, Y) = .Columns(Y).Text
                        Next Y
                        myWS.Cells(z, 1) = .Items.Item(i - 1).SubItems.Item(1).Text
                        myWS.Cells(z, 2) = .Items.Item(i - 1).SubItems.Item(2).Text
                    Next i
                    myWS.Columns(1).ColumnWidth = 50
                    myWS.Columns(2).ColumnWidth = 100
                Case 3
                    For i = 1 To .Items.Count
                        For Y = 1 To 3
                            myWS.Cells(1, Y) = .Columns(Y).Text
                            myWS.Cells(z, 1) = .Items.Item(i - 1).SubItems.Item(1).Text
                            myWS.Cells(z, 2) = .Items.Item(i - 1).SubItems.Item(2).Text
                            myWS.Cells(z, 3) = .Items.Item(i - 1).SubItems.Item(3).Text
                        Next Y
                        z = z + 1
                    Next i
                    myWS.Columns(1).ColumnWidth = 50
                    myWS.Columns(2).ColumnWidth = 100
                    myWS.Columns(3).ColumnWidth = 50
            End Select
        End With
        z = z + 1
        myXL = Nothing
        myWB = Nothing
        myWS = Nothing
    End Sub
End Module
