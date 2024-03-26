Public Class docuDetails
    Private q As New qry
    Public id As String = ""

    Private Sub docuDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        fetchDetails()
    End Sub

    Sub fetchDetails()
        q.fetchDocType(id)
        btnUpdate.Select()
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        With documentTypes
            .Enabled = True
            .loadLV()
            .tbxFilter1.Select()
        End With
        tbxDocName.Text = ""
        tbxDocRem.Text = ""
        id = ""
        rdPrim.Checked = True
        rdRecv.Checked = True
        tbxDocName.Select()
        Me.Dispose()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        validateUpdate()
    End Sub

    Sub validateUpdate()
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

            q.UpdateDocType(id, tbxDocName.Text, tbxDocRem.Text, isRecv, isPrim)
        End If
    End Sub

    Private Sub docuDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closeThisForm()
    End Sub

    Private Sub docuDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            closeThisForm()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class