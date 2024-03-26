Public Class viewLocations
    Private q As New qry
    Private id As String = ""
    Private Sub viewLocations_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If newHome.roleId = 20020 Then
            loadLVadmin()
            newLVFormatadmin()
            Panel4.Visible = True
            q.loadAreasInReportr(ComboBox1)
        Else
            newLVFormat()
            loadLV()
            Panel4.Visible = False
        End If
    End Sub

    Sub newLVFormat()
        With lvLocations
            .Columns.Add("CODE", 0)
            .Columns.Add("LOCATION", 150)
            .Columns.Add("DESCRIPTION", 450)
            .Columns.Add("DATE MODIFIED", 150)
        End With
    End Sub

    Sub newLVFormatadmin()
        With lvLocations
            .Columns.Add("CODE", 0)
            .Columns.Add("LOCATION", 150)
            .Columns.Add("DESCRIPTION", 200)
            .Columns.Add("DATE MODIFIED", 150)
            .Columns.Add("AREA", 200)
        End With
    End Sub

    Sub loadLV()
        q.loadLVLocationsView(lvLocations, tbxFilter, newHome.areaId, newHome.roleId)
        If lvLocations.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub


    Sub loadLVadmin()
        q.loadLVLocationsViewadmin(lvLocations, tbxFilter.Text, ComboBox1.Text)
        If lvLocations.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub



    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        If newHome.roleId = 20020 Then
            loadLVadmin()
        Else
            loadLV()
            id = ""
            enabledBtnUpdate()
            lvLocations.SelectedItems.Clear()
        End If


    End Sub

    Private Sub lvLocations_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvLocations.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvLocations.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Home_Page.Enabled = False
        Me.Enabled = True
        addLocationMain.Show()
        addLocationMain.Focus()
        addLocationMain.Select()
    End Sub


    Sub fetchDatas()
        Dim x As Integer = 0
        x = CInt(lvLocations.SelectedItems(0).Index)
        id = ""
        id = lvLocations.Items(x).SubItems(0).Text
        enabledBtnUpdate()
    End Sub

    Sub enabledBtnUpdate()
        If Not String.IsNullOrWhiteSpace(id) Then
            btnDetails.Enabled = True
        Else
            btnDetails.Enabled = False
        End If
    End Sub

    Private Sub lvLocations_DoubleClick(sender As Object, e As EventArgs) Handles lvLocations.DoubleClick
        With updateLocationMain
            .id = id
            Home_Page.Enabled = False
            Me.Enabled = False
            .Show()
            .Focus()
            .Select()
            .loadDatas()
        End With
    End Sub

    Private Sub lvLocations_Click(sender As Object, e As EventArgs) Handles lvLocations.Click
        fetchDatas()
    End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        With updateLocationMain
            .id = id
            Home_Page.Enabled = False
            Me.Enabled = False
            .Show()
            .Focus()
            .Select()
            .loadDatas()
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub viewLocations_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvLocations.Clear()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            ComboBox1.Items.Clear()
            RadioButton2.Checked = False
            ComboBox1.Visible = False
            loadLVadmin()
        End If


    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            q.loadAreasInReportr(ComboBox1)
            RadioButton1.Checked = False
            ComboBox1.Visible = True
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadLVadmin()
    End Sub

    Private Sub ViewLocations_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class