Public Class TileSetup
    Private n As New qryv3
    Public id As Integer
    Private Sub TileSetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loaddg()
    End Sub


    Public Sub loaddg()
        n.ViewTileSetup(TextBox1.Text)
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Form2.ShowDialog()


    End Sub

    Private Sub Dg1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form3.ShowDialog()



    End Sub

    Private Sub TileSetup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        TextBox1.Text = String.Empty
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loaddg()

        If TextBox1.Text = "" Then
            Button1.Enabled = False
        Else
            Button1.Enabled = True
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs)
        loaddg()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        loaddg()
    End Sub

    Private Sub Dg1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellDoubleClick
        Button1.Enabled = True

        Dim des As String
        Dim mes As String

        If dg1.SelectedRows.Count > 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = dg1.SelectedRows(0)

            ' Check if the cell in the first column of the selected row is not null
            If selectedRow.Cells(0).Value IsNot Nothing Then
                ' Retrieve and display the value in a message box
                id = selectedRow.Cells(0).Value.ToString()
                des = selectedRow.Cells(4).Value.ToString()
                mes = selectedRow.Cells(5).Value.ToString()


                With Form2
                    .id = id
                    .TextBox2.Text = des
                    .TextBox3.Text = mes
                End With

                TextBox1.Text = selectedRow.Cells(1).Value.ToString()
            End If
        End If

    End Sub

    Private Sub TileSetup_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class