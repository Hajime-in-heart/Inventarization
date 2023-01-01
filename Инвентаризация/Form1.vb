Imports System.Data.OleDb
Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvwcolumnsorter = New listviewcolumnsorter
        SysLoad()
    End Sub

    Sub SysLoad()
        With ListView1
            .ListViewItemSorter = lvwcolumnsorter
            .Columns.Clear()
            .Columns.Add("ID")
            .Columns.Add("Номер")
            .Columns.Add("IP")
            .Columns.Add("DNS")
            .Columns.Add("Марка")
            .Columns.Add("S/N")
            .Columns.Add("Устройство")
            .Columns.Add("Наличие")
            .Items.Clear()
        End With
        Try
            request = "Select * From Техника"
            Dim datareader As OleDbDataReader
            Dim command As New OleDbCommand(request, connector)
            Dim kartinka As Integer = 3
            connector.Open()
            datareader = command.ExecuteReader
            While datareader.Read() = True
                If datareader.GetValue(6) = "Видео-фотокамера" Then kartinka = 0
                If datareader.GetValue(6) = "Проектор" Then kartinka = 1
                If datareader.GetValue(6) = "Телефон" Then kartinka = 2
                If datareader.GetValue(6) = "Ноутбук" Then kartinka = 3
                With ListView1
                    .Items.Add(datareader.GetValue(0), kartinka)
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(1))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(2))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(3))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(4))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(5))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(6))
                    .Items.Item(.Items.Count - 1).SubItems.Add(datareader.GetValue(7))
                End With
            End While
            datareader.Close()
            connector.Close()
        Catch ex As Exception
            connector.Close()
            MsgBox(ex.Message)
        End Try
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)
        For Each columnheader In Me.ListView1.Columns
            columnheader.width = -3
        Next
    End Sub

    Private Sub DeleteNote()
        Try
            If Not ListView1.SelectedItems(0).SubItems.Item(7).Text = "" Then
                request = "DELETE FROM [Техника] WHERE ID =" & ListView1.SelectedItems.Item(0).Text & ""
                ChangeDataInTable(request)
                SysLoad()
            End If
        Catch ex As Exception
            MessageBox.Show("Не выбрано устройство" & Chr(13) & Chr(10) & "Выберите устройство...", "Невозможно выполнить операцию...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ОбновитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem.Click
        SysLoad()
    End Sub

    Private Sub ВыходToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыходToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub БольшиеИконкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles БольшиеИконкиToolStripMenuItem.Click
        ListView1.View = View.LargeIcon
    End Sub

    Private Sub ДеталиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ДеталиToolStripMenuItem.Click
        ListView1.View = View.Details
    End Sub

    Private Sub МаленькиеИконкиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles МаленькиеИконкиToolStripMenuItem.Click
        ListView1.View = View.SmallIcon
    End Sub

    Private Sub СписокToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles СписокToolStripMenuItem.Click
        ListView1.View = View.List
    End Sub

    Private Sub НазваниеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles НазваниеToolStripMenuItem.Click
        ListView1.View = View.Tile
    End Sub

    Sub dobavlenie_zap()
        ListView1.ListViewItemSorter = lvwcolumnsorter
        request = "Insert Into [Техника] ([Номер], [IP], [DNS], [Марка], [S/N], [Устройство], [Наличие]) values ('" & номер & "', '" & IP & "','" & DNS & "', '" & Марка & "', '" & SN & "', '" & Устройство & "','" & Наличие & "')"
        ChangeDataInTable(request)
        With ListView1
            .Items.Add("Номер")
            .Items.Item(.Items.Count - 1).SubItems.Add("IP")
            .Items.Item(.Items.Count - 1).SubItems.Add("DNS")
            .Items.Item(.Items.Count - 1).SubItems.Add("Марка")
            .Items.Item(.Items.Count - 1).SubItems.Add("SN")
            .Items.Item(.Items.Count - 1).SubItems.Add("Устройство")
            .Items.Item(.Items.Count - 1).SubItems.Add("Наличие")
        End With
        SysLoad()
        For Each columnheader In ListView1.Columns
            columnheader.width = -3
        Next
    End Sub

    Private Sub ListView1_MouseUp(sender As Object, e As MouseEventArgs) Handles ListView1.MouseUp
        Dim itemindex As Integer
        If e.Button = MouseButtons.Right Then
            itemindex = ListView1.SelectedIndices.Count
            If itemindex = 1 Then
                ContextMenuStrip1.Show(Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip1.Height - 40)
            Else
                ContextMenuStrip2.Show(Me.Location.X + e.X + 10, Location.Y + e.Y + ContextMenuStrip2.Height - 13)
            End If
        End If
    End Sub

    Sub editor_zap()
        request = "UPDATE Техника SET [Номер]='" & номер & "', [IP]='" & IP & "',  [DNS]='" & DNS & "', [Марка]='" & Марка & "', [S/N]='" & SN & "', [Устройство]='" & Устройство & "', [Наличие]='" & Наличие & "' WHERE ([ID] Like '" & ListView1.SelectedItems.Item(0).Text & "')"
        ChangeDataInTable(request)
        With ListView1
            .SelectedItems.Item(0).Text = номер
            .SelectedItems.Item(0).SubItems.Item(1).Text = IP
            .SelectedItems.Item(0).SubItems.Item(2).Text = DNS
            .SelectedItems.Item(0).SubItems.Item(3).Text = Марка
            .SelectedItems.Item(0).SubItems.Item(4).Text = SN
            .SelectedItems.Item(0).SubItems.Item(5).Text = Устройство
            .SelectedItems.Item(0).SubItems.Item(6).Text = Наличие
        End With
        SysLoad()
    End Sub

    Sub выдать()
        Try
            Dim дата As String = Format(Now, "d MMMM yyyy")
            request = "UPDATE Техника SET [Номер]='" & номер & "', [IP]='" & IP & "',  [DNS]='" & DNS & "', [Марка]='" & Марка & "', [S/N]='" & SN & "', [Устройство]='" & Устройство & "', [Наличие]='" & Наличие & "' WHERE ([ID] Like '" & ListView1.SelectedItems.Item(0).Text & "')"
            ChangeDataInTable(request)
            With ListView1
                .SelectedItems.Item(0).Text = номер
                .SelectedItems.Item(0).SubItems.Item(1).Text = IP
                .SelectedItems.Item(0).SubItems.Item(2).Text = DNS
                .SelectedItems.Item(0).SubItems.Item(3).Text = Марка
                .SelectedItems.Item(0).SubItems.Item(4).Text = SN
                .SelectedItems.Item(0).SubItems.Item(5).Text = Устройство
                .SelectedItems.Item(0).SubItems.Item(6).Text = Наличие
            End With
            request = "Insert into [Журнал] ([ФИО], [История], [Дата]) values ('" & Получатель & "', '" & ПК & "', '" & дата & "') "
            ChangeDataInTable(request)
            With ListView1
                .Items.Add("ID")
                .Items.Item(.Items.Count - 1).SubItems.Add("ФИО")
                .Items.Item(.Items.Count - 1).SubItems.Add("История")
                .Items.Item(.Items.Count - 1).SubItems.Add("Дата")
            End With
            request = "Insert into [Журнал выдач] ([ФИО], [История], [IDGurnal]) values ('" & Получатель & "', '" & ПК & "', '" & id & "') "
            ChangeDataInTable(request)
            With ListView1
                .Items.Add("ID")
                .Items.Item(.Items.Count - 1).SubItems.Add("ФИО")
                .Items.Item(.Items.Count - 1).SubItems.Add("История")
                .Items.Item(.Items.Count - 1).SubItems.Add("IDGurnal")
            End With
            SysLoad()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        MessageBox.Show("Устройство: " & Chr(13) & Chr(10) & ПК & Chr(13) & Chr(10) & "Для: " & Получатель & Chr(13) & Chr(10) & "Успешно выдано!!!", "Операция успешно завершена", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Try
            With ListView1
                id = ListView1.SelectedItems.Item(0).Text
                номер = .SelectedItems.Item(0).SubItems.Item(1).Text
                IP = .SelectedItems.Item(0).SubItems.Item(2).Text
                DNS = .SelectedItems.Item(0).SubItems.Item(3).Text
                Марка = .SelectedItems.Item(0).SubItems.Item(4).Text
                SN = .SelectedItems.Item(0).SubItems.Item(5).Text
                Устройство = .SelectedItems.Item(0).SubItems.Item(6).Text
                Наличие = .SelectedItems.Item(0).SubItems.Item(7).Text
            End With
            Form2.Text = "Редактирование записи"
            Form2.Show()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
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
        ListView1.Sort()
    End Sub

    Private Sub УдалитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem1.Click
        Dim mbox As DialogResult = MessageBox.Show("Вы действительно хотите безвозвратно удалить устройство из базы данных?", "Удаление", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation)
        If mbox = DialogResult.No Then Exit Sub
        If mbox = DialogResult.Yes Then
            DeleteNote()
        End If
    End Sub

    Private Sub ПолучателиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПолучателиToolStripMenuItem.Click
        Form3.Text = "Получатели"
        Form3.Show()
    End Sub

    Private Sub ДобавитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem1.Click
        AddNote()
    End Sub

    Private Sub ДобавитьToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ДобавитьToolStripMenuItem.Click
        AddNote()
    End Sub

    Sub AddNote()
        Form2.Text = "Добавление записи"
        Form2.Show()
    End Sub

    Private Sub ВыдатьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ВыдатьToolStripMenuItem.Click
        GiveItem()
    End Sub

    Sub GiveItem()
        Try
            If ListView1.SelectedItems(0).SubItems.Item(7).Text = "" Then
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Не выбрано устройство" & Chr(13) & Chr(10) & "Выберите устройство...", "Невозможно выполнить операцию...", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            If ListView1.SelectedItems.Item(0).SubItems.Item(7).Text = "Да" Then
                id = ListView1.SelectedItems.Item(0).Text
                номер = ListView1.SelectedItems.Item(0).SubItems.Item(1).Text
                IP = ListView1.SelectedItems.Item(0).SubItems.Item(2).Text
                DNS = ListView1.SelectedItems.Item(0).SubItems.Item(3).Text
                Марка = ListView1.SelectedItems.Item(0).SubItems.Item(4).Text
                SN = ListView1.SelectedItems.Item(0).SubItems.Item(5).Text
                Устройство = ListView1.SelectedItems.Item(0).SubItems.Item(6).Text
                Наличие = "Нет"
                ПК = номер & " - " & Марка & " - " & IP & " - " & DNS & " - " & SN & " - " & Устройство
                Form3.Show()
                Form3.Text = "Выберите получателя..."
            Else
                MessageBox.Show("Данное устройство уже было успешно выдано ранее..." & Chr(13) & Chr(10) & "Смотрите журнал выдач...", "Невозможно выполнить операцию...", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ЖурналToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналToolStripMenuItem.Click
        Form5.Text = "Архив выдач..."
        Form5.Show()
    End Sub

    Private Sub ОтчётПоВыданнойТехникиToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОтчётПоВыданнойТехникиToolStripMenuItem.Click
        Form6.Text = "Текущий отчет..."
        Form6.Show()
    End Sub

    Private Sub ЖурналВыдачДанногоУстройстваToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ЖурналВыдачДанногоУстройстваToolStripMenuItem.Click
        With ListView1
            id = .SelectedItems.Item(0).Text
            номер = .SelectedItems.Item(0).SubItems.Item(1).Text
            IP = .SelectedItems.Item(0).SubItems.Item(2).Text
            DNS = .SelectedItems.Item(0).SubItems.Item(3).Text
            Марка = .SelectedItems.Item(0).SubItems.Item(4).Text
            SN = .SelectedItems.Item(0).SubItems.Item(5).Text
            Устройство = .SelectedItems.Item(0).SubItems.Item(6).Text
        End With
        Комп = номер & " - " & Марка & " - " & IP & " - " & DNS & " - " & SN & " - " & Устройство
        Form5.Text = "Отчет по: " & Комп
        Form5.Show()
    End Sub

    Private Sub ТекущийОтчетДанногоУстройстваToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ТекущийОтчетДанногоУстройстваToolStripMenuItem.Click
        With ListView1
            id = .SelectedItems.Item(0).Text
            номер = .SelectedItems.Item(0).SubItems.Item(1).Text
            IP = .SelectedItems.Item(0).SubItems.Item(2).Text
            DNS = .SelectedItems.Item(0).SubItems.Item(3).Text
            Марка = .SelectedItems.Item(0).SubItems.Item(4).Text
            SN = .SelectedItems.Item(0).SubItems.Item(5).Text
            Устройство = .SelectedItems.Item(0).SubItems.Item(6).Text
        End With
        Комп = номер & " - " & Марка & " - " & IP & " - " & DNS & " - " & SN & " - " & Устройство
        Form6.Text = "Отчет текущий по: " & Комп
        Form6.Show()
    End Sub

    Private Sub УдалитьToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles УдалитьToolStripMenuItem.Click
        DeleteNote()
    End Sub

    Private Sub ВыдатьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ВыдатьToolStripMenuItem1.Click
        GiveItem()
    End Sub

    Private Sub ОПрограммеToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОПрограммеToolStripMenuItem.Click
        Form7.Show()
    End Sub

    Private Sub ОАвтореToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ОАвтореToolStripMenuItem.Click
        Form8.Show()
    End Sub

    Private Sub БольшиеИконкиToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles БольшиеИконкиToolStripMenuItem1.Click
        ListView1.View = View.LargeIcon
    End Sub

    Private Sub ДеталиToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ДеталиToolStripMenuItem1.Click
        ListView1.View = View.Details
    End Sub

    Private Sub МаленькиеИконкиToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles МаленькиеИконкиToolStripMenuItem1.Click
        ListView1.View = View.SmallIcon
    End Sub

    Private Sub СписокToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles СписокToolStripMenuItem1.Click
        ListView1.View = View.List
    End Sub

    Private Sub НазваниеToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles НазваниеToolStripMenuItem1.Click
        ListView1.View = View.Tile
    End Sub

    Private Sub ОбновитьToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ОбновитьToolStripMenuItem1.Click
        SysLoad()
    End Sub

    Private Sub ПерезагрузкаToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ПерезагрузкаToolStripMenuItem.Click
        Application.Restart()
    End Sub
End Class