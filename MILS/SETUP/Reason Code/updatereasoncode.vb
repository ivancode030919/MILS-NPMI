Public Class updatereasoncode

    Private q As New qry
    Private Sub Updatereasoncode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = reasoncode.txt1.Text
        TextBox2.Text = reasoncode.txt2.Text
        TextBox3.Text = reasoncode.TextBox2.Text
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        q.updatereasoncode(TextBox1.Text, TextBox2.Text, TextBox3.Text)

        With reasoncode
            .view()
        End With
    End Sub
End Class