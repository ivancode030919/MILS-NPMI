Public Class adjustInventoryRegisterDetails
    Private q As New qry
    Public id As String = ""
    Public encoder As String = ""
    Private Sub adjustInventoryRegisterDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadencoder()
        loadlv()
    End Sub

    Private Sub newLVFormat()
        With lvInvAdjDetails
            .Columns.Add("dummy", 0)
            .Columns.Add("PRODUCT NUMBER", 130, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 400, HorizontalAlignment.Left)
            .Columns.Add("BATCH", 100, HorizontalAlignment.Center)
            .Columns.Add("LOCATION", 100, HorizontalAlignment.Center)
            .Columns.Add("EXPIRATION DATE", 150, HorizontalAlignment.Center)
            .Columns.Add("ON HAND", 90, HorizontalAlignment.Center)
            .Columns.Add("COUNT", 90, HorizontalAlignment.Center)
            .Columns.Add("DIFFERENCE", 150, HorizontalAlignment.Center)
        End With
    End Sub

    Sub loadlv()
        q.fetchDatasInventoryAdjustDetails(lvInvAdjDetails, id)
    End Sub

    Sub loadencoder()
        q.loadforinventorycountencoder(encoder)
    End Sub


    Private Sub adjustInventoryRegisterDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            adjustInventoryRegister.Enabled = True
            Me.Dispose()
            adjustInventoryRegister.Focus()
            adjustInventoryRegister.Select()
        End If
    End Sub

    Private Sub adjustInventoryRegisterDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        adjustInventoryRegister.Enabled = True
        adjustInventoryRegister.Focus()
        adjustInventoryRegister.Select()
    End Sub



    Private Sub lvInvAdjDetails_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvInvAdjDetails.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvInvAdjDetails.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loadencoder()
    End Sub
End Class