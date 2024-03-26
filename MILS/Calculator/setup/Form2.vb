Public Class Form2
    Private n As New qryv3
    Public id As Integer
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Or TextBox3.Text = "" Then
            MsgBox("Please input valid value...")
        Else
            n.UpdateTilesSetup(id, TextBox2.Text, TextBox3.Text)
        End If




    End Sub



    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        TextBox2.Text = ""
    End Sub

    Private Sub TextBox3_Click(sender As Object, e As EventArgs) Handles TextBox3.Click
        TextBox3.Text = ""
    End Sub

    Private Sub Form2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        TextBox3.Text = ""
        TextBox2.Text = ""
    End Sub
End Class