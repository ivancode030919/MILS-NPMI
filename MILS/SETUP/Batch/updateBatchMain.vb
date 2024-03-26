Public Class updateBatchMain
    Private q As New qry
    Public id As String = ""
    Public apprvdBy As String = ""
    Private Sub updateBatchMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadDatas()
    End Sub

    Sub loadDatas()
        q.fetchBatchDatasMain(id)
    End Sub

    Sub validateDatas()
        If String.IsNullOrWhiteSpace(tbxBatchCode.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            Me.Enabled = False
            updateBatchConfirmation.Show()
        End If
    End Sub

    Sub addDatas()
        q.updateBatchMainQry(id, tbxBatchCode.Text, tbxDescrip.Text, newHome.areaId, newHome.userId, apprvdBy)
    End Sub

    Private Sub updateBatchMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        With viewBatch
            .Enabled = True
            .Focus()
            .Select()
            .lvBatches.SelectedItems.Clear()
            .btnDetails.Enabled = False
        End With
    End Sub

    Private Sub updateBatchMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            With viewBatch
                .Enabled = True
                .Focus()
                .Select()
                .lvBatches.SelectedItems.Clear()
                .btnDetails.Enabled = False
            End With
            Me.Dispose()
        Else
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        validateDatas()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class