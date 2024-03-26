Public Class area
    Private q As New qry
    Public id As String = ""
    Private Sub area_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        detailsEnable()
        newLvFormat()
        'loadLvArea()
        loadarea()
        loadallxLvArea()
        If newHome.roleId = "20020" Then
            btnAdd.Enabled = True
            btnDetails.Enabled = True
        End If
    End Sub

    Sub detailsEnable()
        If newHome.roleId = "20020" Then
            If String.IsNullOrWhiteSpace(id) Then
                btnDetails.Enabled = False
            Else
                btnDetails.Enabled = True
            End If
        End If
    End Sub
    Sub loadForm()
        loadallxLvArea()
    End Sub


    Sub loadarea()
        q.loadareainarea(ComboBox1)
    End Sub

    Sub newLvFormat()
        With lvAreas
            .Columns.Add("CODE", 120)
            .Columns.Add("AREA NAME", 200)
            .Columns.Add("DESCRIPTION", 300)
            .Columns.Add("CATEGORY", 134)
        End With
    End Sub

    Sub loadLvArea()
        q.loadArea(lvAreas, tbxFilter, ComboBox1.Text)
        If lvAreas.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub


    Sub loadallxLvArea()
        q.loadallxArea(lvAreas, tbxFilter)
        If lvAreas.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        openAddForm()
    End Sub

    Sub openAddForm()
        Home_Page.Enabled = False
        Me.Enabled = False
        areaAdd.Show()
    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        loadForm()
    End Sub

    Sub fetchId()
        Dim x As Integer = CInt(lvAreas.SelectedItems(0).Index)
        id = ""
        id = lvAreas.Items(x).SubItems(0).Text
        'MsgBox(id)
    End Sub

    Private Sub lvAreas_Click(sender As Object, e As EventArgs) Handles lvAreas.Click
        fetchId()
        detailsEnable()
    End Sub

    'Private Sub lvAreas_DoubleClick(sender As Object, e As EventArgs) Handles lvAreas.DoubleClick
    '    If String.IsNullOrWhiteSpace(id) Then
    '        lblErr.Visible = True
    '        lblErr.Text = "Select valid data to update"
    '    Else
    '        With areaDetails
    '            .id = id
    '            Home_Page.Enabled = False
    '            Me.Enabled = False
    '            .Show()
    '            .fetchId()
    '        End With
    '    End If
    'End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click

        If newHome.roleId = "20020" Then
            If String.IsNullOrWhiteSpace(id) Then
                lblErr.Visible = True
                lblErr.Text = "Select valid data to update"
            Else
                With areaDetails
                    .id = id

                    Me.Enabled = False
                    .Show()
                    .fetchId()
                End With
            End If
        Else

            btnAdd.Enabled = False
            btnDetails.Enabled = False
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) 
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadLvArea()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            ComboBox1.Visible = False
            ComboBox1.Items.Clear()
            loadForm()
            tbxFilter.Text = ""

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            ComboBox1.Items.Clear()
            loadarea()
            lvAreas.Items.Clear()
            RadioButton1.Checked = False
            ComboBox1.Visible = True
            loadLvArea()
            tbxFilter.Text = ""
        End If
    End Sub

    Private Sub area_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvAreas.Clear()
    End Sub

    Private Sub area_HelpRequested(sender As Object, hlpevent As HelpEventArgs) Handles Me.HelpRequested

    End Sub

    Private Sub Area_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class