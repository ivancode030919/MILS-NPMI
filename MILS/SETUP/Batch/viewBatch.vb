Public Class viewBatch
    Private q As New qry
    Private id As String = ""

    Private Sub viewBatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvBatches
            .Columns.Add("ID", 0)
            .Columns.Add("BATCH CODE", 250)
            .Columns.Add("DESCRIPTION", 250)
            .Columns.Add("DATE MODIFIED", 250)
        End With
    End Sub

    Sub loadLV()
        q.loadLVBatchesView(lvBatches, tbxFilter)
        If lvBatches.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        loadLV()
        id = ""
        enabledBtnUpdate()
        lvBatches.SelectedItems.Clear()
    End Sub

    Private Sub lvBatches_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvBatches.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvBatches.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Home_Page.Enabled = False
        Me.Enabled = False
        addBatchMain.Show()
        addBatchMain.Focus()
        addBatchMain.Select()
    End Sub

    Sub fetchDatas()
        Dim x As Integer = 0
        x = CInt(lvBatches.SelectedItems(0).Index)
        id = ""
        id = lvBatches.Items(x).SubItems(0).Text
        enabledBtnUpdate()
    End Sub

    Sub enabledBtnUpdate()
        If Not String.IsNullOrWhiteSpace(id) Then
            btnDetails.Enabled = True
        Else
            btnDetails.Enabled = False
        End If
    End Sub

    Private Sub lvBatches_Click(sender As Object, e As EventArgs) Handles lvBatches.Click
        fetchDatas()
    End Sub

    Private Sub btnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        loadUpdateForm()
    End Sub

    Private Sub lvBatches_DoubleClick(sender As Object, e As EventArgs) Handles lvBatches.DoubleClick
        loadUpdateForm()
    End Sub

    Sub loadUpdateForm()
        With updateBatchMain
            .id = id
            Home_Page.Enabled = False
            Me.Enabled = False
            .Show()
            .loadDatas()
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub viewBatch_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvBatches.Clear()
    End Sub

    Private Sub ViewBatch_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class