Public Class inventoryAutoBuildConfirmation
    Private q As New qry

    Private Sub inventoryAutoBuildConfirmation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        With inventoryAutoBuild
            .Enabled = False
            .Focus()
            .Select()
        End With
    End Sub

    Private Sub inventoryAutoBuildConfirmation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
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
        inventoryAutoBuild.Enabled = True
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

    Private Sub validateUser()
        If String.IsNullOrWhiteSpace(tbxUser.Text) Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            lblError.Visible = True
            lblError.Text = "Fields with * are required."
            Exit Sub
        Else
            lblError.Visible = False
            q.inventoryAutoBuildConfirmationQry(tbxUser.Text, tbxPass.Text, newHome.areaId)
        End If
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        validateUser()
    End Sub
End Class