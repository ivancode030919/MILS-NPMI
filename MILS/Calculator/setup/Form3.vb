Public Class Form3
    Private n As New qryv3
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox3.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Please input valid value...")
        Else
            n.SaveTileSetup(TextBox2.Text, TextBox3.Text)
        End If
    End Sub

    Private Sub Form3_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        TextBox3.Text = ""
        TextBox2.Text = ""

    End Sub
End Class