Public Class itemstatus
    Private n As New qryv3
    Public id As String = ""
    Dim t1 As String


    Private Sub Status_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        n.FetchStatus(TextBox1.Text)

    End Sub

    Private Sub Status_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        inventoryView.Enabled = True
    End Sub


    Public Sub loadstat()
        n.FetchStatus(TextBox1.Text)

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            t1 = "1"
        ElseIf RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            t1 = "2"
        ElseIf RadioButton3.Checked = True Then
            t1 = "3"
        End If


        n.updateGoodStatus(t1, TextBox1.Text)
    End Sub
End Class