Public Class Form4
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Text = "Редактирование записи" Then
            TextBox1.Text = Имя
            TextBox2.Text = Фамилия
            TextBox3.Text = Отчество
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Me.Text = "Добавление записи" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MessageBox.Show("Не все строки заполнены!!!")
            Else
                Фамилия = TextBox1.Text
                Имя = TextBox2.Text
                Отчество = TextBox3.Text
                Form3.Dobavlenie_zap()
                Me.Close()
            End If
        Else
            Фамилия = TextBox1.Text
            Имя = TextBox2.Text
            Отчество = TextBox3.Text
            Form3.Editor_zap()
            Me.Close()
        End If
    End Sub
End Class