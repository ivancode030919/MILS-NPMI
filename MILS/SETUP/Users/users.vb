Public Class users
    Private q As New qry
    Private selectedId As String = ""

    Private Sub users_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()

        detailsEnable()
        loadroles()

        loadLV()

        If newHome.roleId = 20020 Then
            btnAdd.Enabled = True
            btnDetails.Enabled = True
        End If

    End Sub

    Sub newLVFormat()
        With lvUsers
            .Columns.Add("ID", 95)
            .Columns.Add("USER (Firstname Middlename Lastname)", 260)
            .Columns.Add("ROLE", 150)
            .Columns.Add("AREA", 150)
            .Columns.Add("STATUS", 100)
        End With
    End Sub

    Sub loadLV()
        q.loadLVUsers(lvUsers, tbxFN, tbxLN, ComboBox1.Text, newHome.areaId)
        If lvUsers.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
        loadroles()
    End Sub
    Sub loadroles()
        q.loadroles(ComboBox1)
    End Sub
    Private Sub tbxFN_TextChanged(sender As Object, e As EventArgs) Handles tbxFN.TextChanged
        loadLV()
        lvUsers.SelectedItems.Clear()
        selectedId = ""
        detailsEnable()
    End Sub

    Private Sub tbxLN_TextChanged(sender As Object, e As EventArgs) Handles tbxLN.TextChanged
        loadLV()
        lvUsers.SelectedItems.Clear()
        selectedId = ""
        detailsEnable()
    End Sub

    Sub detailsEnable()

        If newHome.roleId = 20020 Then
            If String.IsNullOrWhiteSpace(selectedId) Then
                btnDetails.Enabled = False
            Else
                btnDetails.Enabled = True
            End If
        End If

    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        clickAddUsers()
    End Sub

    Sub clickAddUsers()
        usersAdd.Show()
        Me.Enabled = False
        Home_Page.Enabled = False
        tbxFN.Text = ""
        tbxLN.Text = ""
        selectedId = ""
        lvUsers.SelectedItems.Clear()
        tbxFN.Select()
    End Sub

    Private Sub lvUsers_Click(sender As Object, e As EventArgs) Handles lvUsers.Click
        fetchID()
        detailsEnable()
    End Sub

    Sub fetchID()
        If lvUsers.Items.Count <> 0 Then
            Dim x As Integer = -1
            x = CInt(lvUsers.SelectedItems(0).Index)
            selectedId = ""
            selectedId = lvUsers.Items(x).SubItems(0).Text
            'MsgBox(selectedId)
        End If
    End Sub

    Sub loadFormDetails()


        If newHome.roleId = 20020 Then
            newHome.Enabled = False
            Me.Enabled = False
            With usersDetails
                .Show()
                .id = selectedId
                .fetchData()
                'MsgBox(.id)
            End With
        End If

    End Sub

    Private Sub lvUsers_DoubleClick(sender As Object, e As EventArgs) Handles lvUsers.DoubleClick
        loadFormDetails()
    End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        loadFormDetails()
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadLV()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            ComboBox1.Visible = False
            ComboBox1.Items.Clear()
            loadroles()

            loadLV()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            ComboBox1.Visible = True

        End If
    End Sub

    Private Sub users_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        lvUsers.Clear()
    End Sub

    Private Sub Users_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class