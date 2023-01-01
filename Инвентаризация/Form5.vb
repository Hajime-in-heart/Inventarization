Public Class Form5
    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click
        Sysload()
    End Sub

    Private Sub ЭкспортWordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортWordToolStripMenuItem.Click
        ExportWord(ListView1)
    End Sub

    Private Sub УдалитьИсториюToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьИсториюToolStripMenuItem.Click
        request = "DELETE FROM [Журнал] WHERE ID =" & ListView1.SelectedItems.Item(0).Text & ""
        ChangeDataInTable(request)
        Sysload()
    End Sub

    Private Sub ЭкспортExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортExcelToolStripMenuItem.Click
        ExportExcel(ListView1, 3)
    End Sub

    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvwcolumnsorter = New listviewcolumnsorter()
        Sysload()
    End Sub

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
        Dim itemindex As Integer
        If e.Button = MouseButtons.Right Then
            itemindex = ListView1.SelectedIndices.Count
            If itemindex = 1 Then
                ContextMenuStrip1.Show(Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip1.Height - 40)
            End If
        End If
    End Sub

    Sub Sysload()
        With ListView1
            .ListViewItemSorter = lvwcolumnsorter
            .Columns.Clear()
            .Columns.Add("ID")
            .Columns.Add("ФИО")
            .Columns.Add("История")
            .Columns.Add("Дата и время")
            .Items.Clear()
        End With
        If Me.Text = "Архив выдач..." Then
            request = "Select * From Журнал"
            ReadDataFromTable(request, ListView1, 4)
        End If
        If Me.Text = "Отчет по:  " & ФИО Then
            request = "Select * From Журнал WHERE ФИО = '" & ФИО & "'"
            ReadDataFromTable(request, ListView1, 4)
        End If
        If Me.Text = "Отчет по: " & Комп Then
            request = "Select * From Журнал WHERE [История] = '" & Комп & "'"
            ReadDataFromTable(request, ListView1, 4)
        End If
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In Me.ListView1.Columns
            columnheader.Width = -3
        Next
    End Sub
End Class