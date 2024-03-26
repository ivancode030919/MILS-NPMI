Public Class intraTransferDetails
    Private q As New qry
    Public id As String = ""

    Private Sub intraTransferDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvNewFormat()
        'tbxApprovedBy.Visible = False
        'Label4.Visible = False
        loadLV()
    End Sub

    Sub lvNewFormat()
        With lvIntraTransferDetails
            .Columns.Add("dummy", 0)
            .Columns.Add("PRODUCT NUMBER", 130, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 500, HorizontalAlignment.Left)
            .Columns.Add("BATCH CODE", 100, HorizontalAlignment.Center)
            .Columns.Add("SOURCE LOCATION", 120, HorizontalAlignment.Center)
            .Columns.Add("DESTINATION", 110, HorizontalAlignment.Center)
            .Columns.Add("QUANTITY", 90, HorizontalAlignment.Center)
            .Columns.Add("EXPIRATION DATE", 130, HorizontalAlignment.Center)
        End With
    End Sub

    Sub loadLV()
        q.fetchIntraTransferDetails(lvIntraTransferDetails, id)


    End Sub

    Private Sub intraTransferDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        intraTransferRegister.Enabled = True
        intraTransferRegister.lvIntraTransfer.Select()
    End Sub

    Private Sub intraTransferDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            intraTransferRegister.Enabled = True
            intraTransferRegister.lvIntraTransfer.Select()
            Me.Dispose()
        End If
    End Sub

    Private Sub lvIntraTransferDetails_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvIntraTransferDetails.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvIntraTransferDetails.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
End Class