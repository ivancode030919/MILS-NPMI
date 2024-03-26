Public Class UpdateConfirmation


    Private q As New qry
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
        TextBox1.Text = ""
        TextBox2.Text = ""
        UpdateDes.Show()
        UpdateDes.Enabled = True
        UpdateDes.BringToFront()

    End Sub

    Private Sub UpdateConfirmation_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = UpdateDes.TextBox1.Text
        TextBox2.Text = UpdateDes.destext.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to Update", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        ' Check the user's response
        If result = DialogResult.Yes Then
            ' User clicked Yes

            updatedloadlv()
        Else
            ' User clicked No or closed the message box
            MessageBox.Show("Cancelled.")
            Me.Dispose()
        End If
        Me.Dispose()
        UpdateDes.Show()
        UpdateDes.Enabled = True
        UpdateDes.BringToFront()


        With UpdateDes
            .TextBox1.Text = ""
            .destext.Text = ""
        End With

    End Sub


    Sub updatedloadlv()
        Dim id As String = TextBox1.Text
        q.uodateloadArea(id, TextBox2)

    End Sub

    Private Sub UpdateConfirmation_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        UpdateDes.Enabled = True
    End Sub
End Class