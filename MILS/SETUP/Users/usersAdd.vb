Public Class usersAdd
    Private q As New qry
    Public userAreaId As String = ""
    Public userRoleId As String = ""

    Private Sub usersAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadcbx()
    End Sub

    Sub loadcbx()
        q.loadAreaforUsers(cbxArea)
        q.loadRolesforUsers(cbxRole)
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        validateAdd()
    End Sub

    Sub validateAdd()
        If String.IsNullOrWhiteSpace(tbxFN.Text) Or String.IsNullOrWhiteSpace(tbxLN.Text) _
            Or cbxArea.SelectedIndex = -1 Or cbxRole.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(tbxUserName.Text) _
            Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            MessageBox.Show("Fields with asterisk are required.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            q.addUser(tbxFN.Text, tbxMN.Text, tbxLN.Text, userRoleId, userAreaId, newHome.userId, tbxUserName.Text, tbxPass.Text)
        End If
    End Sub

    Sub closeThisForm()
        With users
            .loadLV()
            .tbxFN.Text = ""
            .tbxLN.Text = ""
        End With
        Home_Page.Enabled = True
        users.Enabled = True
        clearFields()
        Me.Dispose()
    End Sub

    Sub clearFields()
        tbxFN.Text = ""
        tbxMN.Text = ""
        tbxLN.Text = ""
        cbxRole.SelectedIndex = -1
        cbxArea.SelectedIndex = -1
        tbxUserName.Text = ""
        tbxPass.Text = ""
        btnAdd.Select()
    End Sub

    Private Sub usersAdd_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim notNull As Boolean = True
        If Not String.IsNullOrWhiteSpace(tbxFN.Text) Or Not String.IsNullOrWhiteSpace(tbxLN.Text) _
        Or Not cbxArea.SelectedIndex = -1 Or Not cbxRole.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxUserName.Text) _
        Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            notNull = False
        Else
            notNull = True
        End If

        If notNull = False Then
            e.Cancel = True
            If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                'btnClose = True
                e.Cancel = True
            Else
                e.Cancel = False
                Home_Page.Enabled = True
                users.Enabled = True
                With users
                    .loadLV()
                    .tbxFN.Text = ""
                    .tbxLN.Text = ""
                End With
            End If
        Else
            e.Cancel = False
            Home_Page.Enabled = True
            users.Enabled = True
            With users
                .loadLV()
                .tbxFN.Text = ""
                .tbxLN.Text = ""
            End With
        End If
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        userAreaId = q.fetchAreaforUsers(cbxArea.Text)
    End Sub

    Private Sub cbxRole_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxRole.SelectedIndexChanged
        userRoleId = q.fetchRolesforUsers(cbxRole.Text)
    End Sub

    Private Sub usersAdd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'Dim btnClose As Boolean = True
        If e.KeyCode = Keys.Escape Then
            Dim notNull As Boolean = True

            If Not String.IsNullOrWhiteSpace(tbxFN.Text) Or Not String.IsNullOrWhiteSpace(tbxLN.Text) _
                Or Not cbxArea.SelectedIndex = -1 Or Not cbxRole.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxUserName.Text) _
                Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
                notNull = False
            Else
                notNull = True
            End If

            If notNull = False Then
                'e.Cancel = True
                If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                Else
                    closeThisForm()
                    Home_Page.Enabled = True
                    users.Enabled = True
                End If
            Else
                closeThisForm()
                Home_Page.Enabled = True
                users.Enabled = True
            End If

        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim notNull As Boolean = True

        If Not String.IsNullOrWhiteSpace(tbxFN.Text) Or Not String.IsNullOrWhiteSpace(tbxLN.Text) _
            Or Not cbxArea.SelectedIndex = -1 Or Not cbxRole.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxUserName.Text) _
            Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            notNull = False
        Else
            notNull = True
        End If

        If notNull = False Then
            'btnClose = True
            'e.Cancel = True
            If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                'btnClose = True
            Else
                'btnClose = False
                closeThisForm()
                Home_Page.Enabled = True
                users.Enabled = True
            End If
        Else
            'btnClose = False
            closeThisForm()
            Home_Page.Enabled = True
            users.Enabled = True
        End If
    End Sub
End Class