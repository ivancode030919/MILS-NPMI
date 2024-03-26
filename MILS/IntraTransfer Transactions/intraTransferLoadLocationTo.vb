Public Class intraTransferLoadLocationTo
    Private q As New qry
    Public targetRow As String = ""
    Public locId As String = ""
    Private Sub intraTransferLoadLocationTo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvLocations
            .Columns.Add("CODE", 0)
            .Columns.Add("LOCATION", 305)
            .Columns.Add("DESCRIPTION", 450)
        End With
    End Sub

    Sub loadLV()
        q.selectLVLocationsIntraTransfer(lvLocations, tbxFilter, locId)

    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        loadLV()
    End Sub

    Private Sub intraTransferLoadLocationTo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        intraTransfer.Enabled = True
        intraTransfer.Select()
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        intraTransfer.Enabled = True
        Me.Dispose()
        With intraTransfer
            .dgvTransferGoods.ClearSelection()
            .dgvTransferGoods.Select()
        End With
    End Sub

    Sub fetchDatas()
        Try
            Dim x As Integer = -1
            Dim lId As String = ""
            Dim locDes As String = ""
            x = CInt(lvLocations.SelectedItems(0).Index)
            With lvLocations.Items(x)
                lId = .SubItems(0).Text
                locDes = .SubItems(1).Text
            End With
            With intraTransfer.dgvTransferGoods
                .Rows(targetRow).Cells(3).Value = lId
                .Rows(targetRow).Cells(9).Value = locDes
            End With
            closeThisForm()
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Private Sub lvLocations_DoubleClick(sender As Object, e As EventArgs) Handles lvLocations.DoubleClick
        fetchDatas()
    End Sub

    Private Sub intraTransferLoadLocationTo_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        End If
    End Sub

    Private Sub lvLocations_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvLocations.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvLocations.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs)
        fetchDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        closeThisForm()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class