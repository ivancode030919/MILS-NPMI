Public Class updateBatchConfirmation
    Private q As New qry
    Private Sub updateBatchConfirmation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub validateDatas()
        If String.IsNullOrWhiteSpace(tbxUsername.Text) Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.validateUpdateBatchMain(tbxUsername.Text, tbxPass.Text)
        End If
    End Sub

    Private Sub updateBatchConfirmation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        updateBatchMain.Enabled = True
        updateBatchMain.Focus()
        updateBatchMain.Select()
    End Sub

    Private Sub updateBatchConfirmation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxUsername.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
                If (MessageBox.Show("Are you sure to cancel confirmation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    closeThisForm()
                End If
            Else
                closeThisForm()
            End If
        End If
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        If Not String.IsNullOrWhiteSpace(tbxUsername.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            If (MessageBox.Show("Are you sure to cancel confirmation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                closeThisForm()
            End If
        Else
            closeThisForm()
        End If
    End Sub

    Sub closeThisForm()
        updateBatchMain.Enabled = True
        updateBatchMain.Focus()
        updateBatchMain.Select()
        Me.Dispose()
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        validateDatas()
    End Sub
End Class