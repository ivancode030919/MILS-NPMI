Public Class selectProdABLocations
    Private q As New qry
    Public targetRow As Integer = 0
    Private Sub selectProdABLocations_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvLocations
            .Columns.Add("ID", 0)
            .Columns.Add("LOCATION", 150)
            .Columns.Add("DESCRIPTION", 350)
        End With
    End Sub

    Sub loadLV()
        q.loadLV_selectABLocations(lvLocations, tbxLocName, tbxDes)
        If lvLocations.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Private Sub lvLocations_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvLocations.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvLocations.Columns(e.ColumnIndex).Width
    End Sub

    Sub fetchDatas()
        Try
            Dim x As Integer = 0
            x = CInt(lvLocations.SelectedItems(0).Index)

            Dim id As String = ""
            Dim des As String = ""

            id = lvLocations.Items(x).SubItems(0).Text
            des = lvLocations.Items(x).SubItems(1).Text

            With inventoryAutoBuild
                .dgvInventory.Rows(targetRow).Cells(2).Value = id
                .dgvInventory.Rows(targetRow).Cells(5).Value = des
                Home_Page.Enabled = True
                inventoryAutoBuild.Enabled = True
                Me.Dispose()
            End With
        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

    Private Sub lvLocations_DoubleClick(sender As Object, e As EventArgs) Handles lvLocations.DoubleClick
        fetchDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fetchDatas()
    End Sub

    Private Sub selectProdABLocations_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        inventoryAutoBuild.Enabled = True
        inventoryAutoBuild.Focus()
        inventoryAutoBuild.Select()
    End Sub

    Private Sub selectProdABLocations_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            inventoryAutoBuild.Enabled = True
            Me.Dispose()
        End If
    End Sub

    Private Sub tbxDes_TextChanged(sender As Object, e As EventArgs) Handles tbxDes.TextChanged
        loadLV()
    End Sub

    Private Sub tbxLocName_TextChanged(sender As Object, e As EventArgs) Handles tbxLocName.TextChanged
        loadLV()
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Me.Enabled = False
        autoBuildSubFormLocation.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class