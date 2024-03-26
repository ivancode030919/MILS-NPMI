Public Class addNewBatch
    Private q As New qry

    Private Sub addNewBatch_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub closeThisForm()
        recvSelectBatch.Enabled = True
        clearFields()
        With recvSelectBatch
            .tbxFilter.Text = ""
            .tbxFilter.Select()
            .loadLV()
        End With
        Me.Hide()
    End Sub

    Sub clearFields()
        tbxBatchDes.Text = ""
        tbxBatchName.Text = ""
        lblErr.Visible = False
        tbxBatchDes.Select()
    End Sub

    Private Sub addNewBatch_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub

    Private Sub addNewBatch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim btnClose As Boolean
        If e.KeyCode = Keys.Escape Then
            If tbxBatchName.Text <> "" Then
                If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    btnClose = True
                Else
                    btnClose = False
                    closeThisForm()
                End If
            Else
                closeThisForm()
            End If
        Else

        End If
    End Sub


    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxBatchName.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with '*' are required."
            tbxBatchName.Select()
        Else
            lblErr.Visible = False
            q.addBatch(tbxBatchName.Text, tbxBatchDes.Text)
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        validateFields()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim btnClose As Boolean
        If tbxBatchName.Text <> "" Then
            If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                btnClose = True
            Else
                btnClose = False
                closeThisForm()
            End If
        Else
            closeThisForm()
        End If
    End Sub
End Class