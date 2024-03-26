
Public Class EditExpirationdate

    Private q As New qry
    Private dt As New DataTable
    Private columnLVSelected As Integer = 0
    Private isEditedDGV As Boolean = False
    Public approvedID As String = ""

    Private Sub EditExpirationdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadDGV()
    End Sub

    Private Sub loadDGV()
        q.loadexdate(TextBox1.Text, TextBox2.Text, TextBox3.Text)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loadDGV()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        loadDGV()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        loadDGV()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)
        loadDGV()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        q.EditExDate()
    End Sub

    Private Sub Dgexdate_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgexdate.CellDoubleClick, dgexdate.CellClick
        Dim row As Integer = dgexdate.CurrentCell.RowIndex
        Dim ro As DataGridViewRow = dgexdate.Rows(e.RowIndex)
        If e.ColumnIndex = 6 Then
            With editexpiration
                .rowToEdit = row
                .ShowDialog()
                .Focus()
            End With

            With dgexdate
                ro.Cells(7).Value = "1"
            End With
        End If
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Dgexdate_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgexdate.CellValidating

    End Sub

    Private Sub EditExpirationdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress


        'If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
        '    e.Handled = True ' Ignore the input
        '    MessageBox.Show("
        '
        '
        '
        '    ")
        'End If

    End Sub

    Private Sub EditExpirationdate_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class