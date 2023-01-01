Public Class Form6
    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvwcolumnsorter = New listviewcolumnsorter()
        Sysload()
    End Sub

    Private Sub Sysload()
        With ListView1
            .ListViewItemSorter = lvwcolumnsorter
            .Columns.Clear()
            .Columns.Add("ID")
            .Columns.Add("ФИО")
            .Columns.Add("История")
            .Columns.Add("IDGurnal")
            .Items.Clear()
        End With

        If Me.Text = "Текущий отчет..." Then
            request = "Select * From [Журнал выдач]"
            ReadDataFromTable(request, ListView1, 4)
        End If

        If Me.Text = "Отчет текущий по " & ФИО Then
            request = "Select * From [Журнал выдач] WHERE (([Журнал выдач].ФИО) = '" & ФИО & "')"
            ReadDataFromTable(request, ListView1, 4)
        End If

        If Me.Text = "Отчет текущий по: " & Комп Then
            request = "Select * From [Журнал выдач] WHERE [Журнал выдач].[История] = '" & Комп & "'"
            ReadDataFromTable(request, ListView1, 4)
        End If

        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In Me.ListView1.Columns
            columnheader.Width = -3
        Next
    End Sub

    Private Sub ЭкспортWordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортWordToolStripMenuItem.Click
        ExportWord(ListView1)
    End Sub

    Private Sub ЭкспортExcelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЭкспортExcelToolStripMenuItem.Click
        ExportExcel(ListView1, 2)
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click
        Sysload()
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick
        If (e.Column = lvwcolumnsorter.sortcolumn) Then
            If (lvwcolumnsorter.order = SortOrder.Ascending) Then

                lvwcolumnsorter.order = SortOrder.Descending
            Else
                lvwcolumnsorter.order = SortOrder.Ascending
            End If
        Else
            lvwcolumnsorter.sortcolumn = e.Column
            lvwcolumnsorter.order = SortOrder.Ascending
        End If
        Me.ListView1.Sort()
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        ReturnItem()
    End Sub

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
        Dim itemindex As Integer
        If e.Button = MouseButtons.Right Then
            itemindex = ListView1.SelectedIndices.Count
            If Not itemindex = 1 Then
                ContextMenuStrip1.Show(Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip1.Height - 65)
            Else
                ContextMenuStrip2.Show(Me.Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip2.Height)
            End If
        End If
    End Sub

    Private Sub ВНаличииToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВНаличииToolStripMenuItem.Click
        ReturnItem()
    End Sub

    Sub ReturnItem()
        Try
            Dim MBox As DialogResult = MessageBox.Show("Выполнить возврат", "Уведомление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
            If MBox = DialogResult.No Then Exit Sub
            If MBox = DialogResult.Yes Then
                request = "DELETE FROM [Журнал выдач] WHERE ID =" & ListView1.SelectedItems.Item(0).Text & ""
                ChangeDataInTable(request)
                request = "UPDATE Техника SET [Наличие]='Да' WHERE ([ID] Like '" & ListView1.SelectedItems.Item(0).SubItems.Item(3).Text & "')"
                ChangeDataInTable(request)
                Sysload()
                Form1.SysLoad()
                MessageBox.Show("Успешно выполнен возврат", "Операция успешно завершена",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class