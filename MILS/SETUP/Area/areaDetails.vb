Public Class areaDetails
    Private q As New qry
    Public id As String = ""
    Private Sub areaDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fetchId()
    End Sub

    Sub fetchId()
        q.fetchArea(id)
    End Sub

    Sub closeForm()
        Home_Page.Enabled = True
        area.Enabled = True
        With area
            .id = ""
            .detailsEnable()
            .lvAreas.SelectedItems.Clear()
            .tbxFilter.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub areaDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If (MessageBox.Show("Cancel Update?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
            e.Cancel = True
        Else
            e.Cancel = False
            closeForm()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If (MessageBox.Show("Cancel Update?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
        Else
            closeForm()
        End If
    End Sub

    Private Sub areaDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If (MessageBox.Show("Cancel Update?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
            Else
                closeForm()
            End If
        Else

        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrWhiteSpace(tbxAreaName.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with '*' are required."
            tbxAreaName.Select()
        Else
            lblErr.Visible = False
            Dim isInternal As Boolean
            If rdInternal.Checked = True Then
                isInternal = True
            Else
                isInternal = False
            End If
            q.editArea(id, tbxAreaName.Text, tbxAreaDes.Text, isInternal)
        End If
    End Sub
End Class