Public Class addLocationMain
    Private q As New qry

    Private Sub addLocationMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub validateField()
        If String.IsNullOrWhiteSpace(tbxLocName.Text) Or String.IsNullOrWhiteSpace(tbxDescrip.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.addNewLocationMain(tbxLocName.Text, tbxDescrip.Text, newHome.areaId, newHome.userId)
        End If
    End Sub

    Private Sub addLocationMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        viewLocations.Enabled = True
        viewLocations.Focus()
        viewLocations.Select()
    End Sub

    Private Sub addLocationMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxLocName.Text) Or Not String.IsNullOrWhiteSpace(tbxDescrip.Text) Then
                If (MessageBox.Show("Are you sure you want to cancel transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
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

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        validateField()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class