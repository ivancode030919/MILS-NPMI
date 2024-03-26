Public Class selectProdforReserveRelease
    Private q As New qry
    Public targetRow As Integer = -1
    Private id As String
    Private gId As String
    Private desc As String
    Private batch As String
    Private locDes As String
    Private quanti As String
    Private bId As String
    Private lId As String

    Private Sub selectProdforReserveRelease_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newFormatLV()
        loadLV()
    End Sub

    Sub newFormatLV()
        With lvProducts
            .Columns.Add("ID", 0)
            .Columns.Add("PRODUCT NUMBER", 120)
            .Columns.Add("DESCRIPTION", 425)
            .Columns.Add("BATCH CODE", 120)
            .Columns.Add("LOCATION", 120)
            .Columns.Add("QUANTITY", 120)
            .Columns.Add("bId", 0)
            .Columns.Add("lId", 0)
            tbxFilter1.Select()
        End With
    End Sub

    Sub loadLV()
        'q.loadLVProducts(lvProducts, tbxFilter1, tbxFilter2)
        If lvProducts.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        reserveRelease.Enabled = True
        Me.Dispose()
    End Sub

    Private Sub selectProdforReserveRelease_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
            tbxFilter1.Text = ""
            tbxFilter2.Text = ""
            lvProducts.SelectedItems.Clear()
            tbxFilter1.Select()
        End If
    End Sub

    Private Sub selectProdforReserveRelease_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        reserveRelease.Enabled = True
    End Sub

    Private Sub tbxFilter1_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter1.TextChanged
        loadLV()
    End Sub

    Private Sub tbxFilter2_TextChanged(sender As Object, e As EventArgs) Handles tbxFilter2.TextChanged
        loadLV()
    End Sub

    Sub fetchDatas()
        Try
            Dim x As Integer = -1
            x = CInt(lvProducts.SelectedItems(0).Index)

            id = lvProducts.Items(x).SubItems(0).Text
            gId = lvProducts.Items(x).SubItems(1).Text
            desc = lvProducts.Items(x).SubItems(2).Text
            batch = lvProducts.Items(x).SubItems(3).Text
            locDes = lvProducts.Items(x).SubItems(4).Text
            bId = lvProducts.Items(x).SubItems(6).Text
            lId = lvProducts.Items(x).SubItems(7).Text
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        If String.IsNullOrWhiteSpace(id) Then
            'lblErr.Text = "Select valid data..."
            'lblErr.Visible = True

        Else
            With reserveRelease.dgvRels
                .Rows(targetRow).Cells(0).Value = id
                .Rows(targetRow).Cells(1).Value = gId
                .Rows(targetRow).Cells(2).Value = desc
                .Rows(targetRow).Cells(3).Value = batch
                .Rows(targetRow).Cells(4).Value = locDes
                .Rows(targetRow).Cells(6).Value = bId
                .Rows(targetRow).Cells(7).Value = lId

                Home_Page.Enabled = True
                reserveRelease.Enabled = True
                lvProducts.SelectedItems.Clear()
                id = ""
                desc = ""
                gId = ""
                batch = ""
                locDes = ""
                tbxFilter1.Text = ""
                tbxFilter2.Text = ""
                Me.Close()
                .ClearSelection()
            End With
        End If
    End Sub

    Private Sub lvProducts_DoubleClick(sender As Object, e As EventArgs) Handles lvProducts.DoubleClick
        fetchDatas()
    End Sub

    Private Sub lvProducts_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvProducts.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvProducts.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        fetchDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        closeThisForm()
    End Sub
End Class