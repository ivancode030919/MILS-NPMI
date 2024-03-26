Public Class invViewDetails
    Private q As New qry
    Public id As String = ""
    Private Sub invViewDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvProdDes
            .Columns.Add("PRODCUT NUMBER", 125)
            .Columns.Add("DESCRIPTION", 500)
            .Columns.Add("BATCH CODE", 125)
            .Columns.Add("LOCATION", 125)
            .Columns.Add("ON HAND", 125)
        End With
    End Sub

    Sub loadLV()
        q.loadLVProductDetails(lvProdDes, newHome.areaId, id, tbxBatch.Text, tbxLoc.Text)
        If lvProdDes.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub invViewDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        'If e.KeyCode = Keys.Escape() Then
        If e.KeyCode = Keys.Escape Then
            inventoryView.Enabled = True

            With inventoryView
                .loadLV()
                .tbxPDNO.Text = ""
                .tbxProdDes.Text = ""
                .lvInventory.SelectedItems.Clear()
            End With
            Me.Dispose()
            tbxBatch.Text = ""
            tbxLoc.Text = ""
            tbxBatch.Select()
        End If
        'End If
    End Sub

    Private Sub invViewDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        'inventoryView.Enabled = True

        'With inventoryView
        '    .loadLV()
        '    .tbxPDNO.Text = ""
        '    .tbxProdDes.Text = ""
        '    .lvInventory.SelectedItems.Clear()
        'End With
        inventoryView.Enabled = True

        With inventoryView
            .loadLV()
            .tbxPDNO.Text = ""
            .tbxProdDes.Text = ""
            .lvInventory.SelectedItems.Clear()
        End With
        Me.Dispose()
        tbxBatch.Text = ""
        tbxLoc.Text = ""
        tbxBatch.Select()
    End Sub

    Private Sub tbxBatch_TextChanged(sender As Object, e As EventArgs) Handles tbxBatch.TextChanged
        loadLV()
    End Sub

    Private Sub tbxLoc_TextChanged(sender As Object, e As EventArgs) Handles tbxLoc.TextChanged
        loadLV()
    End Sub

    Private Sub lvProdDes_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvProdDes.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvProdDes.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        inventoryView.Enabled = True

        With inventoryView
            .loadLV()
            .tbxPDNO.Text = ""
            .tbxProdDes.Text = ""
            .lvInventory.SelectedItems.Clear()
        End With
        Me.Dispose()
        tbxBatch.Text = ""
        tbxLoc.Text = ""
        tbxBatch.Select()
    End Sub
End Class