Public Class usersDetails
    Public id As String = ""
    Private q As New qry
    Private areaID As String = ""
    Private roleID As String = ""

    Private Sub usersDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadCBX()
    End Sub

    Sub fetchData()
        q.fetchUser(id)
    End Sub

    Sub loadCBX()
        q.loadRolesforUsers(cbxRole)
        q.loadAreaforUsers(cbxArea)
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        areaID = q.fetchAreaforUsers(cbxArea.Text)
    End Sub

    Private Sub cbxRole_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxRole.SelectedIndexChanged
        roleID = q.fetchRolesforUsers(cbxRole.Text)
    End Sub

    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxFN.Text) Or String.IsNullOrWhiteSpace(tbxLN.Text) _
           Or cbxArea.SelectedIndex = -1 Or cbxRole.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(tbxUserName.Text) _
           Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            MessageBox.Show("Fields with asterisk are required.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim status As Boolean = True
            If rdActive.Checked = True Then
                status = True
            ElseIf rdInactive.Checked = True Then
                status = False
            End If
            q.updateUsers(id, tbxFN.Text, tbxMN.Text, tbxLN.Text, roleID, areaID, status, tbxUserName.Text, tbxPass.Text, newHome.userId)
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

    Private Sub usersDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
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
                    .tbxLN.Text = ""
                    .tbxFN.Text = ""
                End With
            End If
        Else
            e.Cancel = False
            Home_Page.Enabled = True
            users.Enabled = True
            With users
                .loadLV()
                .tbxLN.Text = ""
                .tbxFN.Text = ""
            End With
        End If
    End Sub

    Sub clearFields()
        tbxFN.Text = ""
        tbxMN.Text = ""
        tbxLN.Text = ""
        cbxRole.SelectedIndex = -1
        cbxArea.SelectedIndex = -1
        tbxUserName.Text = ""
        tbxPass.Text = ""
        btnUpdate.Select()
    End Sub

    Sub closingForm()
        'Dim btnClose As Boolean = True
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
                Home_Page.Enabled = True
                users.Enabled = True
                closeThisForm()

            End If
        Else
            'btnClose = False
            Home_Page.Enabled = True
            users.Enabled = True
            closeThisForm()

        End If
    End Sub

    Private Sub usersDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
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
        closingForm()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        validateFields()
    End Sub

End Class