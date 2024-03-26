Public Class autoBuildSubFormLocation
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        If MsgBox("Are you sure you want cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Me.Close()
        Else

        End If
    End Sub

    Private Sub autoBuildSubFormLocation_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            If MsgBox("Are you sure you want cancel?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Me.Close()
            Else

            End If
        Else

        End If
    End Sub

    Private Sub autoBuildSubFormLocation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxLocName.Select()
    End Sub

    Private Sub autoBuildSubFormLocation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        With selectProdABLocations
            .Enabled = True
            .Select()
            .Focus()
            .loadLV()
        End With
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If String.IsNullOrWhiteSpace(tbxLocName.Text) Or String.IsNullOrWhiteSpace(tbxLocDes.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            Dim q As New qry
            q.addABProd_LocationForm(tbxLocName.Text, tbxLocDes.Text, newHome.userId, newHome.areaId)
        End If
    End Sub
End Class