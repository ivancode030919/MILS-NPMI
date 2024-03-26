Public Class recvSelectGood
    Private q As New qry
    Private columnLVSelected As Integer = 0
    Public rowToEdit As Integer
    Private pdno As String
    Private prodDes As String
    Public op As Boolean = False
    Public ed As Boolean = False
    Private Sub recvSelectGood_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVformat()
        loadLV()

    End Sub

    Sub newLVformat()
        With lvGoods
            .Columns.Add("PD", 0, HorizontalAlignment.Center)
            .Columns.Add("PD NO.", 150, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 650)
            tbxFilter.Select()
        End With
    End Sub

    Sub loadLV()
        q.loadLVforRecvGoods(lvGoods, tbxFilter, tbxFilter1, columnLVSelected)
        'q.loadLVforRecvGoods(lvGoods, tbxFilter, columnLVSelected)
        'If lvGoods.Items.Count = 0 Then
        '    lblError.Visible = True
        '    lblError.Text = "No data found"
        'Else
        '    lblErr.Visible = False
        'End If
    End Sub


    Private Sub tbxFilter_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter.TextChanged
        loadLV()
    End Sub
    Private Sub tbxFilter1_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter1.TextChanged
        loadLV()
    End Sub

    Private Sub recvSelectGood_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        Else

        End If
    End Sub

    Sub closeThisForm()
        recvGoodsMain.Enabled = True
        Home_Page.Enabled = True
        tbxFilter.Text = ""
        tbxFilter1.Text = "" 'added tbx for description

        With recvGoodsMain
            .dgvRecv.ClearSelection()
            .dgvRecv.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub recvSelectGood_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub


    Sub fetchIDandDes()
        '---------------------------------------------------------------------------------------
        'fetch data to re-order setup
        '---------------------------------------------------------------------------------------
        If op = True Then
            Try
                Dim x As Integer = 0
                x = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(x).SubItems(1).Text
                prodDes = lvGoods.Items(x).SubItems(2).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(pdno) And String.IsNullOrWhiteSpace(prodDes) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(y).SubItems(1).Text
                prodDes = lvGoods.Items(y).SubItems(2).Text
                With reordersetup
                    .dgorder.Rows(rowToEdit).Cells(0).Value = pdno
                    .dgorder.Rows(rowToEdit).Cells(1).Value = prodDes
                    closeThisForm()
                    .dgorder.ClearSelection()
                End With

            End If

            '---------------------------------------------------------------------------------------
            'fetch data to receive deatils for editing purpose
            '---------------------------------------------------------------------------------------
        ElseIf ed = True Then


            Try
                Dim x As Integer = 0
                x = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(x).SubItems(1).Text
                prodDes = lvGoods.Items(x).SubItems(2).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(pdno) And String.IsNullOrWhiteSpace(prodDes) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
                Dim y As Integer = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(y).SubItems(1).Text
                prodDes = lvGoods.Items(y).SubItems(2).Text
                With recvListingDetails
                    .dg1.Rows(rowToEdit).Cells(1).Value = pdno
                    .dg1.Rows(rowToEdit).Cells(2).Value = prodDes
                    closeThisForm()
                    .dg1.ClearSelection()
                End With

            End If

            '---------------------------------------------------------------------------------------
            'fetch data in Receiving
            '---------------------------------------------------------------------------------------

        Else


            Try
                Dim x As Integer = 0
                x = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(x).SubItems(1).Text
                prodDes = lvGoods.Items(x).SubItems(2).Text
            Catch ex As Exception
                MessageBox.Show("Please select valid data.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            If String.IsNullOrWhiteSpace(pdno) And String.IsNullOrWhiteSpace(prodDes) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else

                Dim y As Integer = CInt(lvGoods.SelectedItems(0).Index)

                pdno = lvGoods.Items(y).SubItems(1).Text
                prodDes = lvGoods.Items(y).SubItems(2).Text

                With recvGoodsMain
                    .dgvRecv.Rows(rowToEdit).Cells(0).Value = pdno
                    .dgvRecv.Rows(rowToEdit).Cells(3).Value = prodDes
                    closeThisForm()
                    .dgvRecv.ClearSelection()
                End With



            End If
        End If



    End Sub

    Private Sub lvGoods_DoubleClick(sender As Object, e As EventArgs) Handles lvGoods.DoubleClick
        fetchIDandDes()
    End Sub



    Private Sub lvGoods_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvGoods.ColumnClick
        tbxFilter.Text = ""
        tbxFilter1.Text = "" 'added tbx for description
        tbxFilter.Select()
        tbxFilter1.Select() 'added tbx for description
        columnLVSelected = 0
        columnLVSelected = e.Column.ToString()
        loadLV()
    End Sub

    Private Sub lvGoods_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvGoods.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvGoods.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        closeThisForm()
    End Sub

    Private Sub TbxFilter_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxFilter.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub
End Class