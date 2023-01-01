Public Class Form2
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If Me.Text = "Редактирование записи" Then
            TextBox1.Text = номер
            TextBox2.Text = IP
            TextBox3.Text = DNS
            TextBox4.Text = Марка
            TextBox5.Text = SN
            ComboBox1.Text = Устройство
            ComboBox2.Text = Наличие
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Me.Text = "Добавление записи" Then
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or ComboBox1.Text = "" Or ComboBox2.Text = "" Then
                MessageBox.Show("Не все строки заполнены!!!")
            Else
                номер = TextBox1.Text
                IP = TextBox2.Text
                DNS = TextBox3.Text
                Марка = TextBox4.Text
                SN = TextBox5.Text
                Устройство = ComboBox1.Text
                Наличие = ComboBox2.Text
                Form1.dobavlenie_zap()
                Me.Close()
            End If
        Else
            номер = TextBox1.Text
            IP = TextBox2.Text
            DNS = TextBox3.Text
            Марка = TextBox4.Text
            SN = TextBox5.Text
        End If
    End Sub

End Class