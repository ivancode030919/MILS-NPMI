Public Class rolesDetails
    Private q As New qry
    Public id As String = ""
    Private Sub rolesDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fetchId()
        btnUpdate.Select()
    End Sub

    Sub fetchId()
        q.fetchRoles(id)
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        roles.Enabled = True
        roles.Select()
        roles.Focus()
        With roles
            .loadLV()
            .tbxFilter.Text = ""
            .tbxFilter.Select()
            .id = ""
            .detailsEnable()
        End With
    End Sub

    Private Sub rolesDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        btnUpdate.Select()
        closeThisForm()
        With roles
            .loadLV()
            .tbxFilter.Text = ""
            .tbxFilter.Select()
            .id = ""
            .detailsEnable()
        End With
    End Sub

    Private Sub rolesDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
            clearFields()
            Me.Dispose()
        End If
    End Sub

    Sub clearFields()
        id = ""
        tbxRole.Text = ""
        tbxRoleDes.Text = ""
        btnUpdate.Select()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        validateFields()
    End Sub

    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxRole.Text) Or String.IsNullOrWhiteSpace(tbxRoleDes.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.updateRoles(id, tbxRole.Text, tbxRoleDes.Text)
        End If
    End Sub
End Class