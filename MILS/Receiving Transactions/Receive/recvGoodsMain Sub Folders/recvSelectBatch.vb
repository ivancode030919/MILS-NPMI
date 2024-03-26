Public Class recvSelectBatch
    Private q As New qry
    Private columnLVSelected As Integer = 0
    Public rowToEdit As Integer
    Private batchName As String
    Private batchId As String
    Public ed As Boolean = False
    Private Sub recvSelectBatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVformat()
        loadLV()
    End Sub

    Sub newLVformat()
        With lvBatches
            .Columns.Add("ID", 0)
            .Columns.Add("BATCH CODE", 150)
            .Columns.Add("DESCRIPTION", 350)
            tbxFilter.Select()
        End With
    End Sub

    Sub loadLV()
        q.loadLVforRecvBatches(lvBatches, tbxFilter)
        afterDataBind()
    End Sub

    Sub afterDataBind()
        If lvBatches.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub



    Sub closeThisForm()
        recvGoodsMain.Enabled = True
        Home_Page.Enabled = True
        tbxFilter.Text = ""
        lblError.Visible = False
        With recvGoodsMain
            .dgvRecv.ClearSelection()
            .dgvRecv.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        q.loadLVforRecvBatches(lvBatches, tbxFilter)
        afterDataBind()
    End Sub

    Private Sub recvSelectBatch_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub

    Private Sub recvSelectBatch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        Else

        End If
    End Sub

    Sub fetchIDandDes()

        If ed = True Then
            Try
                Dim x As Integer = 0
                x = CInt(lvBatches.SelectedItems(0).Index)

                batchId = lvBatches.Items(x).SubItems(0).Text
                batchName = lvBatches.Items(x).SubItems(1).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(batchId) And String.IsNullOrWhiteSpace(batchName) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvBatches.SelectedItems(0).Index)


                batchName = lvBatches.Items(y).SubItems(1).Text
                With recvListingDetails
                    .dg1.Rows(rowToEdit).Cells(3).Value = batchName
                    .dg1.Rows(rowToEdit).Cells(8).Value = batchId
                    closeThisForm()
                    .dg1.ClearSelection()
                End With
            End If
        Else
            Try
                Dim x As Integer = 0
                x = CInt(lvBatches.SelectedItems(0).Index)

                batchId = lvBatches.Items(x).SubItems(0).Text
                batchName = lvBatches.Items(x).SubItems(1).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(batchId) And String.IsNullOrWhiteSpace(batchName) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvBatches.SelectedItems(0).Index)

                batchId = lvBatches.Items(y).SubItems(0).Text
                batchName = lvBatches.Items(y).SubItems(1).Text
                With recvGoodsMain
                    .dgvRecv.Rows(rowToEdit).Cells(1).Value = batchId
                    .dgvRecv.Rows(rowToEdit).Cells(4).Value = batchName
                    closeThisForm()
                    .dgvRecv.ClearSelection()
                End With
            End If
        End If


    End Sub

    'Private Sub btnOk_Click(sender As Object, e As EventArgs) Handles btnOk.Click
    '    fetchIDandDes()
    'End Sub

    Private Sub lvBatches_DoubleClick(sender As Object, e As EventArgs) Handles lvBatches.DoubleClick
        fetchIDandDes()
    End Sub

    Private Sub lvBatches_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvBatches.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvBatches.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        closeThisForm()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        addBatchMain.ShowDialog()
    End Sub
End Class