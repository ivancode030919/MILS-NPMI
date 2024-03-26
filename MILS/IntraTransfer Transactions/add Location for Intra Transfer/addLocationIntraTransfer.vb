Public Class addLocationIntraTransfer
    Private q As New qry

    Private Sub addLocationIntraTransfer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        With intraTransferLoadLocationTo
            .Enabled = True
            .Focus()
            .Select()
            .loadLV()
        End With
    End Sub

    Private Sub addLocationIntraTransfer_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxLocName.Text) Or Not String.IsNullOrWhiteSpace(tbxLocDes.Text) Then
                If MsgBox("Are you sure you want to cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    Me.Close()
                Else

                End If
            Else
                Me.Close()
            End If
        ElseIf e.KeyCode = Keys.Escape Then
            validateDatas()
        End If
    End Sub

    Private Sub addLocationIntraTransfer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxLocName.Select()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If Not String.IsNullOrWhiteSpace(tbxLocName.Text) Or Not String.IsNullOrWhiteSpace(tbxLocDes.Text) Then
            If MsgBox("Are you sure you want to cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Me.Close()
            Else

            End If
        Else
            Me.Close()
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        validateDatas()
    End Sub

    Sub validateDatas()
        If String.IsNullOrWhiteSpace(tbxLocName.Text) Or String.IsNullOrWhiteSpace(tbxLocDes.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            q.addLocation_IntraTransfer(tbxLocName.Text, tbxLocDes.Text, newHome.areaId, newHome.userId)
        End If
    End Sub
End Class