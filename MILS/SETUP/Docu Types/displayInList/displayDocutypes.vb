Public Class displayDocutypes
    Private categ As String = ""
    Private type As String = ""

    Private Sub displayDocutypes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Private Sub newLVFormat()
        With lvDoctypes
            .Columns.Add("CODE", 0)
            .Columns.Add("DOCUMENT CODE", 150)
            .Columns.Add("DESCRIPTION", 305)
            .Columns.Add("CATEGORY", 150)
            .Columns.Add("TYPE", 150)
        End With
    End Sub

    Sub loadLV()
        Dim q As New qry
        'MsgBox(tbxLocName.Text & "%" + tbxDescrip.Text & "%" + categ + type)
        q.loadDisplayDocuTypesLV(lvDoctypes, tbxLocName.Text & "%", tbxDescrip.Text & "%", categ, type)
        If lvDoctypes.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
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

    Private Sub lvDoctypes_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvDoctypes.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvDoctypes.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub tbxCancel_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub tbxLocName_TextChanged(sender As Object, e As EventArgs) Handles tbxLocName.TextChanged
        loadLV()
    End Sub

    Private Sub tbxDescrip_TextChanged(sender As Object, e As EventArgs) Handles tbxDescrip.TextChanged
        loadLV()
    End Sub

    Private Sub DisplayDocutypes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvDoctypes.Clear()
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub DisplayDocutypes_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class