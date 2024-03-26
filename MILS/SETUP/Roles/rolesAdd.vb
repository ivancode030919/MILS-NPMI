Public Class rolesAdd
    Private q As New qry
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        validateFields()
    End Sub

    Sub clearFields()
        tbxRole.Text = ""
        tbxRoleDes.Text = ""
        btnAdd.Select()
        roles.tbxFilter.Text = ""
    End Sub

    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxRole.Text) Or String.IsNullOrWhiteSpace(tbxRoleDes.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.addRoles(tbxRole.Text, tbxRoleDes.Text)
        End If
    End Sub

    Private Sub rolesAdd_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not String.IsNullOrWhiteSpace(tbxRole.Text) Or Not String.IsNullOrWhiteSpace(tbxRoleDes.Text) Then
            If (MessageBox.Show("Cancel Add?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                e.Cancel = True
            Else
                Home_Page.Enabled = True
                roles.Enabled = True
                e.Cancel = False
                roles.Select()
                roles.Focus()
                roles.tbxFilter.Text = ""
            End If
        Else
            Home_Page.Enabled = True
            roles.Enabled = True
            e.Cancel = False
            roles.Select()
            roles.Focus()
            roles.tbxFilter.Text = ""
        End If
    End Sub

    Private Sub rolesAdd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxRole.Text) Or Not String.IsNullOrWhiteSpace(tbxRoleDes.Text) Then
                If (MessageBox.Show("Cancel Add?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    'e.Cancel = True
                Else
                    closeThisForm()
                    'e.Cancel = False
                End If
            Else
                'e.Cancel = False
                closeThisForm()
            End If
        End If
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        roles.Enabled = True
        Me.Dispose()
        clearFields()
        With roles
            .loadLV()
            .id = ""
            .detailsEnable()
        End With
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If Not String.IsNullOrWhiteSpace(tbxRole.Text) Or Not String.IsNullOrWhiteSpace(tbxRoleDes.Text) Then
            If (MessageBox.Show("Cancel Add?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                'e.Cancel = True
            Else
                closeThisForm()
                'e.Cancel = False
            End If
        Else
            'e.Cancel = False
            closeThisForm()
        End If
    End Sub

End Class