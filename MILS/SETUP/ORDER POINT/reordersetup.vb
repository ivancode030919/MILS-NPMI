Public Class reordersetup

    Private q As New qry
    Private Sub reordersetup_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dgdisplaycolumn()
    End Sub


    Private Sub dgdisplaycolumn()
        With dgorder
            .Columns.Add("0", "DUMMY")
            .Columns.Add("1", "PRODUCT NUMBER")
            .Columns.Add("2", "DESCRIPTION")
            .Columns.Add("3", "ANNUAL TOTAL SALES")
            .Columns.Add("4", "LEAD TIME")
            .Columns.Add("5", "LEAD TIME ALLOWANCE")

            .Columns(0).Visible = False
            .Columns(1).Width = 150
            .Columns(2).Width = 620
            .Columns(3).Width = 150
            .Columns(4).Width = 150
            .Columns(5).Width = 150

            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(5).ReadOnly = True

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .ColumnHeadersDefaultCellStyle.Font = New Font(dgorder.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)

            Next

        End With
    End Sub

    Private Sub dgorder_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgorder.CellDoubleClick
        Dim row As Integer = dgorder.CurrentCell.RowIndex
        Dim ro As DataGridViewRow = dgorder.Rows(e.RowIndex)
        If e.ColumnIndex = 1 Then
            With SELECTGOODS
                .loadLV()
                .rowToEdit = row
                .ShowDialog()
                .Focus()
            End With
        End If
    End Sub



    Private Sub Reordersetup_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        dgorder.Rows.Clear()
        dgorder.Columns.Clear()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ImportAnnualSales.ShowDialog()
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click

        If btnAdd.Text = "Record" Then
            q.setupOrderpoint()
        ElseIf btnAdd.Text = "New Entry" Then

            dgorder.Columns.Clear()
            dgorder.Rows.Clear()


        End If






    End Sub

    Private Sub T1_TextChanged(sender As Object, e As EventArgs) Handles t1.TextChanged
        dgdisplaycolumn()
    End Sub

    Private Sub T2_TextChanged(sender As Object, e As EventArgs) Handles t2.TextChanged
        dgdisplaycolumn()
    End Sub

    Private Sub Dgorder_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgorder.CellValidating
        Dim row As DataGridViewRow = dgorder.Rows(e.RowIndex)
        Dim countedQty As Double = 0
        Dim onhandQty As Double = 0
        Dim differenceQty As Double = 0
        Try
            If e.ColumnIndex = 4 Then
                With dgorder

                    'countedQty = row.Cells(3).Value.ToString
                    onhandQty = e.FormattedValue
                    differenceQty = onhandQty * 0.25

                    row.Cells(5).Value = differenceQty.ToString

                End With
            End If
        Catch ex As Exception
            MessageBox.Show("Please Input Lead Time!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub T1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles t1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub Reordersetup_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class