Public Class goodsRegisterDetails
    Private q As New qry
    Public id As String = ""

    Private Sub goodsRegisterDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        tbxApproved.Visible = False
        Label4.Visible = False
        fetchData(id)
    End Sub

    Sub newLVFormat()
        With lvDetails
            .Columns.Add("PRODUCT NUMBER", 150)
            .Columns.Add("DESCRIPTION", 605)
        End With
    End Sub

    Sub fetchData(id As String)
        q.loadGoodsRegisterDetails(lvDetails, id)
    End Sub

    Private Sub goodsRegisterDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        goodsRegister.Enabled = True
        goodsRegister.Select()
    End Sub

    Private Sub goodsRegisterDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            goodsRegister.Enabled = True
            Me.Dispose()
            goodsRegister.Select()
        End If
    End Sub

    Private Sub lvDetails_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvDetails.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvDetails.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub bntClose_Click(sender As Object, e As EventArgs) Handles bntClose.Click
        Home_Page.Enabled = True
        goodsRegister.Enabled = True
        Me.Dispose()
        goodsRegister.Select()
    End Sub
End Class