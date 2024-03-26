Public Class AgingSetup
    Private n As New qryv3

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Mrupload.ShowDialog()
    End Sub

    Private Sub AgingSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        display()
    End Sub

    Private Sub display()
        n.Aging(t1.Text, t2.Text)
    End Sub

    Private Sub T1_TextChanged(sender As Object, e As EventArgs) Handles t1.TextChanged
        display()
    End Sub

    Private Sub T2_TextChanged(sender As Object, e As EventArgs) Handles t2.TextChanged
        display()
    End Sub

    Private Sub AgingSetup_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class