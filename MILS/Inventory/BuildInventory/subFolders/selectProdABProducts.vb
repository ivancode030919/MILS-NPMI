Public Class selectProdABProducts
    Private q As New qry
    Private columnLVSelected As Integer = 0
    Private id As String
    Private des As String
    Public targetRow As Integer = 0

    Private Sub selectProdABProducts_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvProducts
            .Columns.Add("PD NO.", 0)
            .Columns.Add("PD NO.", 150, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 589)
            'tbxProdId.Select()
        End With
    End Sub

    Sub loadLV()
        q.loadLV_selectABProducts(lvProducts, tbxProdId, tbxDes)
        If lvProducts.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
        'afterDataBinds()
    End Sub

    Sub afterDataBinds()
        If lvProducts.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub
    'tbxProdId the textbox of Product Number
    Private Sub tbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadLV()
    End Sub
    'tbxDes the textbox of Description (Product Description)
    Private Sub tbxDes_TextChanged(sender As Object, e As EventArgs) Handles tbxDes.TextChanged
        loadLV()
    End Sub

    Private Sub selectProdABProducts_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        End If
        '    Home_Page.Enabled = True
        '    inventoryAutoBuild.Enabled = True
        '    Me.Dispose()
        'End If
    End Sub
    Sub closeThisForm()
        inventoryAutoBuild.Enabled = True
        Home_Page.Enabled = True
        tbxProdId.Text = ""
        tbxDes.Text = ""
        lblError.Visible = False
        With inventoryAutoBuild
            .dgvInventory.ClearSelection()
            .dgvInventory.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub selectProdABProducts_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        inventoryAutoBuild.Enabled = True
        inventoryAutoBuild.Focus()
        inventoryAutoBuild.Select()
        closeThisForm()
    End Sub

    Sub fetchDatas()
        Try

            Dim x As Integer = 0
            x = CInt(lvProducts.SelectedItems(0).Index)

            Dim id As String = ""
            Dim des As String = ""

            id = lvProducts.Items(x).SubItems(1).Text
            des = lvProducts.Items(x).SubItems(2).Text

            With inventoryAutoBuild
                .dgvInventory.Rows(targetRow).Cells(0).Value = id
                .dgvInventory.Rows(targetRow).Cells(3).Value = des
                'closeThisForm()
                Home_Page.Enabled = True
                '.dgvInventory.ClearSelection()
                inventoryAutoBuild.Enabled = True
                Me.Dispose()
            End With
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'If String.IsNullOrWhiteSpace(id) And String.IsNullOrWhiteSpace(des) Then

        'Else
        '    Dim y As Integer = CInt(lvProducts.SelectedItems(0).Text)
        '    id = lvProducts.Items(y).SubItems(0).Text
        '    des = lvProducts.Items(y).SubItems(1).Text



        'End If
    End Sub
    Private Sub lvProducts_DoubleClick(sender As Object, e As EventArgs) Handles lvProducts.DoubleClick
        fetchDatas()
    End Sub

    Private Sub lvProducts_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvProducts.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvProducts.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub ButtonOk_Click(sender As Object, e As EventArgs) Handles ButtonOk.Click
        fetchDatas()
    End Sub

    Private Sub lvProducts_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles lvProducts.ColumnClick
        tbxProdId.Text = ""
        tbxDes.Text = ""
        tbxProdId.Select()
        tbxDes.Select()
        columnLVSelected = 0
        columnLVSelected = e.Column.ToString()
        loadLV()
    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        Me.Close()
    End Sub

    Private Sub TbxProdId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdId.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '    ")
        'End If
    End Sub
End Class