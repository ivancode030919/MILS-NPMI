Public Class documentTypes
    Private q As New qry
    Public id As String = ""
    Private categ As String = ""
    Private type As String = ""
    Private Sub documentTypes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()

        If newHome.roleId = "20020" Then
            btnDetails.Enabled = True
            btnAdd.Enabled = True
        Else
            btnDetails.Enabled = False
            btnAdd.Enabled = False
        End If
    End Sub

    Sub newLVFormat()
        With lvDocTypes
            .Columns.Add(" CODE", 0)
            .Columns.Add("DOCUMENT CODE", 150)
            .Columns.Add("DESCRIPTION", 305)
            .Columns.Add("CATEGORY", 150)
            .Columns.Add("TYPE", 150)
        End With
    End Sub

    Sub loadLV()
        'q.loadLVDocTypes(lvDocTypes, tbxFilter1, tbxFilter2)
        q.loadDisplayDocuTypesLV(lvDocTypes, tbxFilter1.Text & "%", tbxFilter2.Text & "%", categ, type)
        If lvDocTypes.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
        id = ""
        detailsEnable()
        lvDocTypes.SelectedItems.Clear()
    End Sub

    Sub detailsEnable()

        If newHome.roleId = "20020" Then
            If String.IsNullOrWhiteSpace(id) Then
                btnDetails.Enabled = False
            Else
                btnDetails.Enabled = True
            End If

        Else
            btnDetails.Enabled = False

        End If


    End Sub

    Private Sub tbxFilter1_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter1.TextChanged
        loadLV()
    End Sub

    Private Sub tbxFilter2_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter2.TextChanged
        loadLV()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Me.Enabled = False
        Home_Page.Enabled = False
        docuAdd.Show()
        docuAdd.Select()
        docuAdd.Focus()
    End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        If String.IsNullOrWhiteSpace(id) Then
            lblErr.Visible = True
            lblErr.Text = "Select valid data to update"
        Else
            With docuDetails
                .id = id
                Home_Page.Enabled = False
                Me.Enabled = False
                .Show()
                .fetchDetails()
            End With
        End If
    End Sub

    Private Sub lvDocTypes_Click(sender As Object, e As EventArgs) Handles lvDocTypes.Click
        fetchId()
    End Sub

    Sub fetchId()
        Dim x As Integer = CInt(lvDocTypes.SelectedItems(0).Index)
        id = ""
        id = lvDocTypes.Items(x).SubItems(0).Text
        detailsEnable()
    End Sub

    'Private Sub lvDocTypes_DoubleClick(sender As Object, e As EventArgs) Handles lvDocTypes.DoubleClick
    '    If String.IsNullOrWhiteSpace(id) Then
    '        lblErr.Visible = True
    '        lblErr.Text = "Select valid data to update"
    '    Else
    '        With docuDetails
    '            .id = id
    '            Home_Page.Enabled = False
    '            Me.Enabled = False
    '            .Show()
    '            .fetchDetails()
    '        End With
    '    End If
    'End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub rdAllCateg_CheckedChanged(sender As Object, e As EventArgs) Handles rdAllCateg.CheckedChanged
        cbxCategHide()
    End Sub

    Private Sub rdSelectCategory_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCategory.CheckedChanged
        cbxCategHide()
    End Sub

    Private Sub cbxCategHide()
        With cbxCategory
            If rdAllCateg.Checked = True Then
                categ = "%"
                .Visible = False
                .SelectedIndex = -1
            ElseIf rdSelectCategory.Checked = True Then
                .Visible = True
                .SelectedIndex = 0
            End If
        End With
        loadLV()
    End Sub

    Private Sub cbxCategory_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxCategory.SelectedIndexChanged
        If cbxCategory.Text = "Receiving" Then
            categ = "1%"
        ElseIf cbxCategory.Text = "Releasing" Then
            categ = "0%"
        End If
        loadLV()
    End Sub

    Private Sub cbxType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxType.SelectedIndexChanged
        If cbxType.Text = "Primary" Then
            type = "1%"
        ElseIf cbxType.Text = "Secondary" Then
            type = "0%"
        End If
        loadLV()
    End Sub

    Private Sub rdAllType_CheckedChanged(sender As Object, e As EventArgs) Handles rdAllType.CheckedChanged
        cbxTypeHide()
    End Sub

    Private Sub rdSelectType_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectType.CheckedChanged
        cbxTypeHide()
    End Sub
    Private Sub cbxTypeHide()
        With cbxType
            If rdAllType.Checked = True Then
                type = "%"
                .Visible = False
                .SelectedIndex = -1
            ElseIf rdSelectType.Checked = True Then
                .Visible = True
                .SelectedIndex = 0
            End If
        End With
        loadLV()
    End Sub

    Private Sub documentTypes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvDocTypes.Clear()
    End Sub

    Private Sub DocumentTypes_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class