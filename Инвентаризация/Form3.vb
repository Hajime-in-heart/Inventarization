Public Class Form3
    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvwcolumnsorter = New listviewcolumnsorter()
        SysLoad()
    End Sub

    Sub SysLoad()
        With ListView1
            .ListViewItemSorter = lvwcolumnsorter
            .Columns.Clear()
            .Columns.Add("ID")
            .Columns.Add("Фамилия")
            .Columns.Add("Имя")
            .Columns.Add("Отчество")
            .Items.Clear()
        End With
        request = "Select * From Получатели"
        ReadDataFromTable(request, ListView1, 4)
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In Me.ListView1.Columns
            columnheader.width = -3
        Next
    End Sub

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
        Dim ItemIndex As Integer
        If e.Button = MouseButtons.Right Then
            ItemIndex = ListView1.SelectedIndices.Count
            If ItemIndex = 1 Then
                ContextMenuStrip2.Show(Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip2.Height - 38)
            Else
                ContextMenuStrip1.Show(Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip1.Height - 65)
            End If
        End If
    End Sub

    Sub Dobavlenie_zap()
        lvwcolumnsorter = New listviewcolumnsorter()
        ListView1.ListViewItemSorter = lvwcolumnsorter
        request = "Insert Into Получатели ([Фамилия], [Имя], [Отчество]) values ('" & Фамилия & "', '" & Имя & "' , '" & Отчество & "')"
        ChangeDataInTable(request)
        With ListView1
            .Items.Add("Фамилия")
            .Items.Item(.Items.Count - 1).SubItems.Add("Имя")
            .Items.Item(.Items.Count - 1).SubItems.Add("Отчество")
        End With
        SysLoad()
    End Sub

    Sub Editor_zap()
        request = "UPDATE Получатели SET [Фамилия]='" & Фамилия & "' [Имя] = '" & Имя & "' [Отчество] = '" & Отчество & "' WHERE ([ID] Like'" & ListView1.SelectedItems.Item(0).Text & "')"
        ChangeDataInTable(request)
        With ListView1
            .SelectedItems.Item(0).Text = Фамилия
            .SelectedItems.Item(0).SubItems.Item(1).Text = Имя
            .SelectedItems.Item(0).SubItems.Item(2).Text = Отчество
        End With
        SysLoad()
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click
        SysLoad()
    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click
        Form4.Text = "Добавление записи"
        Form4.Show()
    End Sub

    Private Sub ЖурналToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналToolStripMenuItem.Click
        ФИО = ListView1.SelectedItems.Item(0).SubItems.Item(1).Text & " " & ListView1.SelectedItems.Item(0).SubItems.Item(2).Text & " " & ListView1.SelectedItems.Item(0).SubItems.Item(3).Text
        Form5.Text = "Отчет по: " & ФИО
        Form5.Show()
    End Sub

    Private Sub ТекущееСостояниеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТекущееСостояниеToolStripMenuItem.Click
        ФИО = ListView1.SelectedItems.Item(0).SubItems.Item(1).Text & " " & ListView1.SelectedItems.Item(0).SubItems.Item(2).Text & " " & ListView1.SelectedItems.Item(0).SubItems.Item(3).Text
        Form6.Text = "Отчет текущий по " & ФИО
        Form6.Show()
    End Sub

    Private Sub ListView1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles ListView1.MouseDoubleClick
        If Me.Text = "Получатели" Then
            Try
                With ListView1
                    id = .SelectedItems.Item(0).Text
                    Фамилия = .SelectedItems.Item(0).SubItems.Item(1).Text
                    Имя = .SelectedItems.Item(0).SubItems.Item(2).Text
                    Отчество = .SelectedItems.Item(0).SubItems.Item(3).Text
                End With
                Form4.Text = "Редактирование записи"
                Form4.Show()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            With ListView1
                Получатель = .SelectedItems.Item(0).SubItems.Item(1).Text & " " &
            .SelectedItems.Item(0).SubItems.Item(2).Text & " " &
            .SelectedItems.Item(0).SubItems.Item(3).Text
            End With
            Form1.выдать()
            Close()
        End If
    End Sub

    Private Sub УдалитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem.Click
        Dim MBox As DialogResult = MessageBox.Show("Вы действительно хотите безвозвратно удалить получателя из базы данных?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If MBox = DialogResult.No Then Exit Sub
        If MBox = DialogResult.Yes Then
            request = "DELETE FROM Получатели WHERE ID =" & ListView1.SelectedItems.Item(0).Text & ""
            ChangeDataInTable(request)
            ListView1.SelectedItems.Item(0).Remove()
            SysLoad()
        End If
    End Sub
End Class