Public Class consolidatedInvDetails
    Public prodId As String = ""
    Public prodArea As String = ""
    Private q As New qry

    Private Sub consolidatedInvDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Private Sub newLVFormat()
        With lvConsoInvDetails
            .Columns.Add("dummy", 0, HorizontalAlignment.Center)
            .Columns.Add("PRODUCT NUMBER", 170, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 452)
            .Columns.Add("BATCH CODE", 150)
            .Columns.Add("LOCATION", 150)
            .Columns.Add("ON HAND", 130)
            .Columns.Add("EXPIRATION DATE", 150)
        End With

    End Sub

    Sub loadLV()
        q.loadLVforConsoInvDetails(lvConsoInvDetails, prodId, prodArea, tbxBatchCode.Text, tbxLoc.Text)
        If lvConsoInvDetails.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub consolidatedInvDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        consolidatedInventoryList.Enabled = True
        consolidatedInventoryList.Focus()
        consolidatedInventoryList.Select()
    End Sub

    Private Sub consolidatedInvDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then

            consolidatedInventoryList.Enabled = True
            consolidatedInventoryList.Focus()
            consolidatedInventoryList.Select()
            Me.Dispose()
        End If
    End Sub

    Private Sub lvConsoInvDetails_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvConsoInvDetails.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvConsoInvDetails.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)

        consolidatedInventoryList.Enabled = True
        consolidatedInventoryList.Focus()
        consolidatedInventoryList.Select()
        Me.Dispose()
    End Sub

    Private Sub tbxBatchCode_TextChanged(sender As Object, e As EventArgs) Handles tbxBatchCode.TextChanged
        loadLV()
    End Sub

    Private Sub tbxLoc_TextChanged(sender As Object, e As EventArgs) Handles tbxLoc.TextChanged
        loadLV()
    End Sub
End Class