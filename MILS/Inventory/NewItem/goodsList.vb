Public Class goodsList
    Private q As New qry

    Private Sub goodsList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvGoods
            .Columns.Add("PRODUCT NUMBER", 200)
            .Columns.Add("DESCRIPTION", 538)
        End With
    End Sub

    Sub loadLV()
        q.loadGoodsListLV(lvGoods, tbxProdId, tbxDescription)
        If lvGoods.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub tbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadLV()
    End Sub

    Private Sub tbxDescription_TextChanged(sender As Object, e As EventArgs) Handles tbxDescription.TextChanged
        loadLV()
    End Sub

    Private Sub lvGoods_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvGoods.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvGoods.Columns(e.ColumnIndex).Width
    End Sub



    Private Sub goodsList_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvGoods.Columns.Clear()
    End Sub

    Private Sub TbxProdId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdId.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub GoodsList_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class