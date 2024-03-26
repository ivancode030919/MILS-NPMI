Public Class MainVendor

    Private q As New qryv3
    Private Sub MainVendor_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loaddg()
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click

        With AddandUpdateVendor
            .Text = "Add"
            .SimpleButton1.Text = "Save"
            .TextBox1.Text = String.Empty
            .ShowDialog()
        End With

    End Sub

    Public Sub loaddg()
        q.loadLVRecvDetail(dg1, TextBox1.Text)
    End Sub

    Private Sub MainVendor_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub dg1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellDoubleClick
        Dim row As Integer = dg1.CurrentCell.RowIndex

        With AddandUpdateVendor
            .id = dg1.Rows(row).Cells(0).Value
            .TextBox1.Text = dg1.Rows(row).Cells(1).Value
            .Text = "Update"
            .SimpleButton1.Text = "Update"
            .ShowDialog()
        End With

    End Sub
End Class