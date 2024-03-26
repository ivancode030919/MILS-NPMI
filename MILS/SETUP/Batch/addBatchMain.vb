Public Class addBatchMain
    Private Sub addBatchMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxBatchCode.Text) Then
            lblErr.Text = "Fields with * are required."
            lblErr.Visible = True
        Else
            lblErr.Visible = False
            Dim q As New qry
            q.addNewBatchMain(tbxBatchCode.Text, tbxDescrip.Text, newHome.userId)
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        validateFields()
    End Sub

    Private Sub addBatchMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        viewBatch.Enabled = True
        viewBatch.Focus()
        viewBatch.Select()
    End Sub

    Private Sub addBatchMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxBatchCode.Text) Then
                If (MessageBox.Show("Are you sure you want to close module?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    Exit Sub
                Else
                    closeThisForm()
                End If
            Else
                closeThisForm()
            End If
        End If
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        viewBatch.Enabled = True
        viewBatch.Focus()
        viewBatch.Select()
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class