Public Class AddandUpdateVendor

    Private q As New qryv3
    Public id As Integer

    Private Sub AddandUpdateVendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        If SimpleButton1.Text = "Save" Then
            q.SaveVendor(TextBox1.Text)
        ElseIf SimpleButton1.Text = "Update" Then
            q.updateGoodStatus(id, TextBox1.Text)
        End If

    End Sub

    Private Sub AddandUpdateVendor_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        TextBox1.Text = String.Empty
    End Sub

    Private Sub AddandUpdateVendor_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class