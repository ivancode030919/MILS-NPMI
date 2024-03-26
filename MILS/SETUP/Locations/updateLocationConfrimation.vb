Public Class updateLocationConfrimation
    Private q As New qry

    Private Sub updateLocationConfrimation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        With updateLocationMain
            .Enabled = True
            .Focus()
            .Select()
        End With
    End Sub



    Sub validateDatas()
        If String.IsNullOrWhiteSpace(tbxUser.Text) Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.validateUpdateLocationMain(tbxUser.Text, tbxPass.Text)
        End If
    End Sub

    Private Sub btnConfirm_Click(sender As Object, e As EventArgs) Handles btnConfirm.Click
        validateDatas()
    End Sub

    Sub closeThisForm()
        updateLocationMain.Enabled = True
        updateLocationMain.Focus()
        updateLocationMain.Select()
        Me.Dispose()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            If (MessageBox.Show("Are you sure to cancel confirmation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                closeThisForm()
            End If
        Else
            closeThisForm()
        End If
    End Sub

    Private Sub updateLocationConfrimation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
                If (MessageBox.Show("Are you sure to cancel confirmation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    closeThisForm()
                End If
            Else
                closeThisForm()
            End If
        End If
    End Sub
End Class