Public Class areaAdd
    Private q As New qry
    Private Sub areaAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxAreaName.Select()
    End Sub

    Sub clearFields()
        tbxAreaDes.Text = ""
        tbxAreaName.Text = ""
        rdInternal.Checked = True
        tbxAreaName.Select()
    End Sub

    Sub closeForm()
        Home_Page.Enabled = True
        area.Enabled = True
        With area
            .tbxFilter.Text = ""
            .tbxFilter.Focus()
            .tbxFilter.Select()
        End With
        clearFields()
        Me.Dispose()
    End Sub

    Sub closingPrompt()
        Dim btnClose As Boolean = True
        If (MessageBox.Show("Cancel Add?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
            btnClose = True
        Else
            btnClose = False
            With area
                .loadForm()
            End With
            closeForm()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        closingPrompt()
    End Sub

    Private Sub areaAdd_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closingPrompt()
    End Sub

    Private Sub areaAdd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closingPrompt()
        Else

        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrWhiteSpace(tbxAreaName.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with '*' are required!"
            tbxAreaName.Select()
        Else
            lblErr.Visible = False
            Dim isInternal As Boolean
            If rdInternal.Checked = True Then
                isInternal = True
            Else
                isInternal = False
            End If
            q.addArea(tbxAreaName.Text, tbxAreaDes.Text, isInternal)
        End If
    End Sub

End Class