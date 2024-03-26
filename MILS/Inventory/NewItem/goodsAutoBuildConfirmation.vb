Public Class goodsAutoBuildConfirmation
    Private q As New qry

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        validateDatas()
    End Sub

    Sub validateDatas()
        If String.IsNullOrWhiteSpace(tbxUser.Text) Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            lblError.Visible = True
            lblError.Text = "Fields with * are required."
        Else
            lblError.Visible = False
            q.goodsAutoBuildConfirmationQry(tbxUser.Text, tbxPass.Text)
        End If
    End Sub

    Private Sub goodsAutoBuildConfirmation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            If (MessageBox.Show("Are you sure to cancel confimation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                closeThisForm()
                'Me.Dispose()
            End If
        Else
            closeThisForm()
            'Me.Dispose()
        End If
    End Sub

    Private Sub goodsAutoBuildConfirmation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
                If (MessageBox.Show("Are you sure to cancel confimation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    closeThisForm()
                    Me.Dispose()
                End If
            Else
                closeThisForm()
                Me.Dispose()
            End If
        End If
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        goodsAutoBuild.Enabled = True
        goodsAutoBuild.Focus()
        goodsAutoBuild.Select()
    End Sub

    Private Sub btnCanel_Click(sender As Object, e As EventArgs) Handles btnCanel.Click
        If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            If (MessageBox.Show("Are you sure to cancel confimation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                closeThisForm()
                Me.Dispose()
            End If
        Else
            closeThisForm()
            Me.Dispose()
        End If
    End Sub
End Class