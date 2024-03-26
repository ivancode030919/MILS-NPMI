Public Class selectProdABBatch
    Private q As New qry
    Public targetRow As Integer = 0

    Private Sub selectProdABBatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvBatch
            .Columns.Add("ID", 0)
            .Columns.Add("BATCH", 150)
            .Columns.Add("DESCRIPTION", 350)
        End With
    End Sub

    Sub loadLV()
        q.loadLV_selectABBatch(lvBatch, tbxBatchCode, tbxDes)
        If lvBatch.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Private Sub selectProdABBatch_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        inventoryAutoBuild.Enabled = True
        inventoryAutoBuild.Focus()
        inventoryAutoBuild.Select()
    End Sub

    Private Sub selectProdABBatch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            inventoryAutoBuild.Enabled = True
            Me.Dispose()
        End If
    End Sub

    Private Sub tbxBatchCode_TextChanged(sender As Object, e As EventArgs) Handles tbxBatchCode.TextChanged
        loadLV()
    End Sub

    Private Sub tbxDes_TextChanged(sender As Object, e As EventArgs) Handles tbxDes.TextChanged
        loadLV()
    End Sub

    Sub fetchDatas()
        Try
            Dim x As Integer = 0
            x = CInt(lvBatch.SelectedItems(0).Index)
            Dim id As String = ""
            Dim des As String = ""

            id = lvBatch.Items(x).SubItems(0).Text
            des = lvBatch.Items(x).SubItems(1).Text

            With inventoryAutoBuild
                .dgvInventory.Rows(targetRow).Cells(1).Value = id
                .dgvInventory.Rows(targetRow).Cells(4).Value = des
                Home_Page.Enabled = True
                inventoryAutoBuild.Enabled = True
                Me.Dispose()
            End With
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub lvBatch_DoubleClick(sender As Object, e As EventArgs) Handles lvBatch.DoubleClick
        fetchDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fetchDatas()
    End Sub

    Private Sub lvBatch_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvBatch.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvBatch.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Me.Enabled = False
        autoBuildSubFormBatch.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class