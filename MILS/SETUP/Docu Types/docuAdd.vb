Public Class docuAdd
    Private q As New qry
    Private Sub docuAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub closethisform()
        Home_Page.Enabled = True
        With documentTypes
            .Enabled = True
            .tbxFilter1.Text = ""
            .tbxFilter1.Select()
            .tbxFilter2.Text = ""
            .loadLV()
            .lvDocTypes.SelectedItems.Clear()
            .id = ""
            .detailsEnable()
            .Select()
            .Focus()

        End With
        tbxDocName.Text = ""
        tbxDocName.Select()
        tbxDocRem.Text = ""
        rdRecv.Checked = True
        rdPrim.Checked = True
        lblErr.Visible = False
        Me.Dispose()
    End Sub

    Private Sub docuAdd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If tbxDocName.Text <> "" Then
                If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    closethisform()
                End If
            Else
                closethisform()
            End If
        End If
    End Sub

    Private Sub docuAdd_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If tbxDocName.Text <> "" Then
            If (MessageBox.Show("Cancel?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                e.Cancel = True
            Else
                e.Cancel = False
                closethisform()
            End If
        Else
            e.Cancel = False
            closethisform()
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrWhiteSpace(tbxDocName.Text) Or String.IsNullOrWhiteSpace(tbxDocRem.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            Dim isRecv As Boolean
            Dim isPrim As Boolean

            If rdRecv.Checked = True Then
                isRecv = True
            Else
                isRecv = False
            End If

            If rdPrim.Checked = True Then
                isPrim = True
            Else
                isPrim = False
            End If

            q.addNewDocType(tbxDocName.Text, tbxDocRem.Text, isRecv, isPrim)
            'MsgBox(isRecv.ToString + " ," + isPrim.ToString)
            tbxDocName.Text = ""
            tbxDocRem.Text = ""
            rdPrim.Checked = True
            rdRecv.Checked = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class