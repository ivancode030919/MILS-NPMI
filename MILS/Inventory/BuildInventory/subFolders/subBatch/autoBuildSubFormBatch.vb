Public Class autoBuildSubFormBatch

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MsgBox("Are you sure you want cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Me.Close()
        Else

        End If
    End Sub

    Private Sub autoBuildSubFormBatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxBatchName.Select()
    End Sub

    Private Sub autoBuildSubFormBatch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If MsgBox("Are you sure you want cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Me.Close()
            Else

            End If
        Else

        End If
    End Sub

    Private Sub autoBuildSubFormBatch_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        With selectProdABBatch
            .Enabled = True
            .Focus()
            .Select()
            .loadLV()
        End With
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrWhiteSpace(tbxBatchName.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            Dim q As New qry
            q.addABProd_BatchForm(tbxBatchName.Text, tbxBatchDes.Text, newHome.userId)
        End If
    End Sub
End Class