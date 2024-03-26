Public Class reorderingpoint
    Private q As New qry
    Private Sub Reorderingpoint_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        view()
        lvcolumn()
    End Sub

    Private Sub lvcolumn()

        With lvop
            .Columns.Add("dummy", 0, HorizontalAlignment.Center)
            .Columns.Add("PRODUCT NUMBER", 150, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 530)
            .Columns.Add("RE-ORDER QUANTITY", 170, HorizontalAlignment.Center)
            .Columns.Add("QUANTITY ON HAND", 170, HorizontalAlignment.Center)
            .Columns.Add("REMARKS", 100, HorizontalAlignment.Center)
            .Columns.Add("QUANTITY TO BUY", 170, HorizontalAlignment.Center)
        End With
    End Sub


    Public Sub view()
        q.loadorderpoint(lvop, txt1.Text, txt2.Text)
    End Sub

    Private Sub reorderingpoint_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        With lvop
            .Clear()
        End With
        txt1.Text = ""
        txt2.Text = ""
        lvop.Clear()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) 
        reordersetup.ShowDialog()
        With reordersetup

        End With
    End Sub

    Private Sub Lvop_DoubleClick(sender As Object, e As EventArgs) Handles lvop.DoubleClick
        If lvop.Items.Count > 0 Then
            Dim x As Integer = CInt(lvop.SelectedItems(0).Index)
            Dim id As String = ""
            id = lvop.Items(x).SubItems(1).Text

            With Orderpoitdetails
                .prodId = ""
                .prodId = id
                .loadodd()
                .ShowDialog()
            End With
        Else
        End If
    End Sub

    Private Sub Txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged
        view()
    End Sub

    Private Sub Txt2_TextChanged(sender As Object, e As EventArgs) Handles txt2.TextChanged
        view()
    End Sub

    Private Sub Txt1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub Reorderingpoint_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class