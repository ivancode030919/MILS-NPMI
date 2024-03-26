Public Class Orderpoitdetails
    Public prodId As String = ""
    Private q As New qry
    Private Sub Orderpoitdetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lvcolumn()
        loadodd()
    End Sub

    Private Sub lvcolumn()

        With opdetail
            .Columns.Add("dummy", 0, HorizontalAlignment.Center)
            .Columns.Add("PRODUCT NUMBER", 150, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 480)
            .Columns.Add("AREA", 200, HorizontalAlignment.Center)
            .Columns.Add("BATCH", 100, HorizontalAlignment.Center)
            .Columns.Add("LOCATION", 100, HorizontalAlignment.Center)
            .Columns.Add("QUANTITY", 100, HorizontalAlignment.Center)
            .Columns.Add("EXPIRATION DATE", 150, HorizontalAlignment.Center)
        End With
    End Sub



    Sub loadodd()
        q.fetchorderpointdetails(opdetail, prodId)
    End Sub

    Private Sub Orderpoitdetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        With opdetail
            .Clear()
        End With
    End Sub

    Private Sub Orderpoitdetails_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class