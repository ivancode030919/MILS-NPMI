Public Class addNewLocation
    Private q As New qry
    Private Sub addNewLocation_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub closeThisForm()
        recvSelectLocations.Enabled = True
        clearFields()
        With recvSelectLocations
            .tbxFilter.Text = ""
            .tbxFilter.Select()
            .loadLV()
        End With
        Me.Hide()
    End Sub

    Sub clearFields()
        tbxLocDes.Text = ""
        tbxLocName.Text = ""
        lblErr.Visible = False
        tbxLocName.Select()
    End Sub

    Private Sub addNewLocation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub

    Private Sub addNewLocation_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim btnClose As Boolean
        If e.KeyCode = Keys.Escape Then
            If tbxLocName.Text <> "" Then
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

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim btnClose As Boolean
        If tbxLocName.Text <> "" Then
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

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        validateFields()
    End Sub

    Sub validateFields()
        If String.IsNullOrWhiteSpace(tbxLocName.Text) Or String.IsNullOrWhiteSpace(tbxLocDes.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with '*' are required."
        Else
            lblErr.Visible = False
            q.addLocations(tbxLocName.Text, tbxLocDes.Text)
        End If
    End Sub
End Class