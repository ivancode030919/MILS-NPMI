Public Class intrTransferLoadProducts
    Public targetRow As Integer = -1
    Private columnLVSelected As Integer = 0
    Private q As New qry
    Private pId As String = ""
    Private gId As String = ""
    Private aId As String = ""
    Private bId As String = ""
    Private lIdF As String = ""
    Private gDes As String = ""
    Private locDes As String = ""
    Private batchDes As String = ""
    Private exdate As String = ""
    Private qty As String = ""

    Private Sub intrTransferLoadProducts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvProducts
            .Columns.Add("id", 0) '0
            .Columns.Add("PD NO.", 160, HorizontalAlignment.Center) '1
            .Columns.Add("DESCRIPTION", 500) '2
            .Columns.Add("areaID", 0) '3
            .Columns.Add("bId", 0) '4
            .Columns.Add("locId", 0) '5
            .Columns.Add("BATCH CODE", 110) '6
            .Columns.Add("LOCATION", 110) '7
            .Columns.Add("QUANTITY", 100) '8
            .Columns.Add("EXPIRATION DATE", 160) '8
        End With
    End Sub

    Sub loadLV()
        q.loadProductforTransfer(lvProducts, tbxProdId, tbxDescrip, newHome.areaId)
        'If lvProducts.Items.Count = 0 Then
        '    lblErr.Visible = True
        '    lblErr.Text = "No data found..."
        'Else
        '    lblErr.Visible = False
        'End If
        afterDataBind()
    End Sub
    Sub afterDataBind()
        If lvProducts.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub
    'tbxProdId the textbox of Product Number
    Private Sub tbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadLV()
    End Sub
    'tbxDes the textbox of Description (Product Description)
    Private Sub tbxDescrip_TextChanged(sender As Object, e As EventArgs) Handles tbxDescrip.TextChanged
        loadLV()
    End Sub

    Private Sub lvProducts_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvProducts.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvProducts.Columns(e.ColumnIndex).Width
    End Sub
    Sub closeThisForm()
        Home_Page.Enabled = True
        intraTransfer.Enabled = True
        tbxProdId.Text = ""
        tbxDescrip.Text = ""
        With intraTransfer
            .dgvTransferGoods.ClearSelection()
            .dgvTransferGoods.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub intrTransferLoadProducts_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        intraTransfer.Enabled = True
        With intraTransfer
            .dgvTransferGoods.ClearSelection()
        End With
    End Sub

    Private Sub intrTransferLoadProducts_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
            tbxProdId.Text = ""
            tbxDescrip.Text = ""
            lvProducts.SelectedItems.Clear()
            tbxProdId.Select()
            tbxDescrip.Select()
        End If
    End Sub

    Sub fetchDatas()
        Try
            Dim x As Integer = -1
            x = CInt(lvProducts.SelectedItems(0).Index)

            pId = lvProducts.Items(x).SubItems(0).Text
            gId = lvProducts.Items(x).SubItems(1).Text
            gDes = lvProducts.Items(x).SubItems(2).Text
            aId = lvProducts.Items(x).SubItems(3).Text
            bId = lvProducts.Items(x).SubItems(4).Text
            lIdF = lvProducts.Items(x).SubItems(5).Text
            batchDes = lvProducts.Items(x).SubItems(6).Text
            locDes = lvProducts.Items(x).SubItems(7).Text
            exdate = lvProducts.Items(x).SubItems(9).Text
            qty = lvProducts.Items(x).SubItems(8).Text
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        If String.IsNullOrWhiteSpace(gId) Then
        Else
            'With lvProducts.Items(x)
            '    pId = .SubItems(0).Text
            '    gId = .SubItems(1).Text
            '    gDes = .SubItems(2).Text
            '    aId = .SubItems(3).Text
            '    bId = .SubItems(4).Text
            '    lIdF = .SubItems(5).Text
            '    batchDes = .SubItems(6).Text
            '    locDes = .SubItems(7).Text
            'End With
            With intraTransfer.dgvTransferGoods
                .Rows(targetRow).Cells(0).Value = pId
                .Rows(targetRow).Cells(5).Value = gId
                .Rows(targetRow).Cells(6).Value = gDes
                .Rows(targetRow).Cells(2).Value = lIdF
                .Rows(targetRow).Cells(4).Value = aId
                .Rows(targetRow).Cells(8).Value = locDes
                .Rows(targetRow).Cells(7).Value = batchDes
                .Rows(targetRow).Cells(1).Value = bId
                .Rows(targetRow).Cells(11).Value = exdate
                .Rows(targetRow).Cells(12).Value = qty
                Home_Page.Enabled = True
                intraTransfer.Enabled = True
                lvProducts.SelectedItems.Clear()
                pId = ""
                gId = ""
                gDes = ""
                lIdF = ""
                aId = ""
                locDes = ""
                batchDes = ""
                bId = ""
                'tbxProdId = ""
                'tbxDescrip = ""
                Me.Dispose()
                .ClearSelection()
            End With
            closeThisForm()

            Exit Sub
        End If

    End Sub

    Private Sub lvProducts_DoubleClick(sender As Object, e As EventArgs) Handles lvProducts.DoubleClick
        fetchDatas()
    End Sub

    Private Sub btnOk_Click(sender As Object, e As EventArgs)
        fetchDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        closeThisForm()
    End Sub

    Private Sub TbxProdId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdId.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub
End Class