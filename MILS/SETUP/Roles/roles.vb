Public Class roles
    Private q As New qry
    Public id As String = ""

    Private Sub roles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()

    End Sub

    Sub newLVFormat()
        With lvRoles
            .Columns.Add("CODE", 120)
            .Columns.Add("ROLE", 200)
            .Columns.Add("DESCRIPTION", 450)
        End With
    End Sub

    Sub loadLV()
        q.loadLVRoles(lvRoles, tbxFilter)
        detailsEnable()
        If lvRoles.Items.Count = 0 Then

        Else

        End If
    End Sub

    Sub detailsEnable()
        If String.IsNullOrWhiteSpace(id) Then
            btnDetails.Enabled = False
        Else
            btnDetails.Enabled = True
        End If
    End Sub

    Sub fetchId()
        Dim x As Integer = -1
        x = CInt(lvRoles.SelectedItems(0).Index)
        id = ""
        id = lvRoles.Items(x).SubItems(0).Text
    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        lvRoles.SelectedItems.Clear()
        id = ""
        loadLV()
    End Sub

    Private Sub lvRoles_Click(sender As Object, e As EventArgs) Handles lvRoles.Click
        fetchId()
        detailsEnable()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        openAddForm()
    End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        updateDetails()
    End Sub

    Sub openAddForm()
        Home_Page.Enabled = False
        Me.Enabled = False
        rolesAdd.Show()
    End Sub

    Sub updateDetails()
        If String.IsNullOrWhiteSpace(id) Then

        Else
            With rolesDetails
                .id = id
                Home_Page.Enabled = False
                Me.Enabled = False
                .Show()
                .fetchId()
                .Select()
            End With
        End If
    End Sub

    Private Sub lvRoles_DoubleClick(sender As Object, e As EventArgs) Handles lvRoles.DoubleClick
        updateDetails()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub roles_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvRoles.Clear()
    End Sub

    Private Sub Roles_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class