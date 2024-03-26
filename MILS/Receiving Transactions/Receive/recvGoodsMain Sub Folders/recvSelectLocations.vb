Public Class recvSelectLocations
    Private q As New qry
    Private columnLVSelected As Integer = 1
    Public rowToEdit As Integer
    Private locName As String
    Private locId As String
    Public ed As Boolean = False
    Private Sub recvSelectLocations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxFilter.Select()
        newLVformat()
        loadLV()
    End Sub

    Sub newLVformat()
        With lvLocations
            .Columns.Add("ID", 0)
            .Columns.Add("LOCATION", 150)
            .Columns.Add("DESCRIPTION", 350)
            tbxFilter.Select()
        End With
    End Sub

    Sub loadLV()
        q.loadLVRecvLocations(lvLocations, tbxFilter)

    End Sub




    Sub closeThisForm()
        recvGoodsMain.Enabled = True
        Home_Page.Enabled = True
        tbxFilter.Text = ""

        With recvGoodsMain
            .dgvRecv.ClearSelection()
            .dgvRecv.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub recvSelectLocations_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub

    Private Sub recvSelectLocations_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        Else

        End If
    End Sub

    Sub fetchIDandDes()
        If ed = True Then
            Try
                Dim x As Integer = CInt(lvLocations.SelectedItems(0).Index)

                locId = lvLocations.Items(x).SubItems(0).Text
                locName = lvLocations.Items(x).SubItems(1).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(locId) And String.IsNullOrWhiteSpace(locName) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvLocations.SelectedItems(0).Index)

                locId = lvLocations.Items(y).SubItems(0).Text
                locName = lvLocations.Items(y).SubItems(1).Text


                With recvListingDetails
                    .dg1.Rows(rowToEdit).Cells(5).Value = locName
                    .dg1.Rows(rowToEdit).Cells(2).Value = locId
                    closeThisForm()
                    .dg1.ClearSelection()
                End With

            End If



        Else

            Try
                Dim x As Integer = CInt(lvLocations.SelectedItems(0).Index)

                locId = lvLocations.Items(x).SubItems(0).Text
                locName = lvLocations.Items(x).SubItems(1).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(locId) And String.IsNullOrWhiteSpace(locName) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvLocations.SelectedItems(0).Index)

                locId = lvLocations.Items(y).SubItems(0).Text
                locName = lvLocations.Items(y).SubItems(1).Text
                With recvGoodsMain
                    .dgvRecv.Rows(rowToEdit).Cells(2).Value = locId
                    .dgvRecv.Rows(rowToEdit).Cells(5).Value = locName
                    closeThisForm()
                    .dgvRecv.ClearSelection()
                End With
            End If

        End If

    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs) 
        fetchIDandDes()
    End Sub

    Private Sub lvLocations_DoubleClick(sender As Object, e As EventArgs) Handles lvLocations.DoubleClick
        fetchIDandDes()
    End Sub

    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        loadLV()
    End Sub

    Private Sub lvLocations_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvLocations.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvLocations.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        closeThisForm()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        addLocationMain.ShowDialog()
    End Sub
End Class