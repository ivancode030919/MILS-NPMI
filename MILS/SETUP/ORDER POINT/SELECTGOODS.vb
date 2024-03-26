Public Class SELECTGOODS

    Private q As New qry
    Private pdno As String
    Private prodDes As String
    Private sales As String
    Private lead As String
    Private leada As String
    Public rowToEdit As Integer
    Private Sub SELECTGOODS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadLV()
        newLVformat()
    End Sub


    Sub newLVformat()
        With listb
            .Columns.Add("PD", 0, HorizontalAlignment.Center)
            .Columns.Add("PD NO.", 100, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 650)
            .Columns.Add("Annual Total Sales", 100, HorizontalAlignment.Center)
            .Columns.Add("LEAD TIME", 100, HorizontalAlignment.Center)
            .Columns.Add("LEAD TIME ALLOWANCE", 100, HorizontalAlignment.Center)
        End With
    End Sub

    Sub loadLV()
        q.loadorderpointsetup(listb, t1.Text, t2.Text)
    End Sub



    Sub fetchIDandDes()

        Try
            Dim x As Integer = 0
            x = CInt(listb.SelectedItems(0).Index)

            pdno = listb.Items(x).SubItems(1).Text
            prodDes = listb.Items(x).SubItems(2).Text
            sales = listb.Items(x).SubItems(3).Text
            lead = listb.Items(x).SubItems(4).Text
            leada = listb.Items(x).SubItems(5).Text

        Catch ex As Exception
            MessageBox.Show("Please select valid data.", "Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


        If String.IsNullOrWhiteSpace(pdno) And String.IsNullOrWhiteSpace(prodDes) Then
                'lblError.Text = "Select valid data!"
                'lblError.Visible = True
            Else
            Dim y As Integer = CInt(listb.SelectedItems(0).Index)

            pdno = listb.Items(y).SubItems(1).Text
            prodDes = listb.Items(y).SubItems(2).Text
            sales = listb.Items(y).SubItems(3).Text
            lead = listb.Items(y).SubItems(4).Text
            leada = listb.Items(y).SubItems(5).Text
            With reordersetup

                .dgorder.Rows(rowToEdit).Cells(1).Value = pdno
                .dgorder.Rows(rowToEdit).Cells(2).Value = prodDes
                .dgorder.Rows(rowToEdit).Cells(3).Value = sales
                .dgorder.Rows(rowToEdit).Cells(4).Value = lead
                .dgorder.Rows(rowToEdit).Cells(5).Value = leada

                .dgorder.ClearSelection()
                closeThisForm()
            End With

        End If



    End Sub

    Private Sub Listb_DoubleClick(sender As Object, e As EventArgs) Handles listb.DoubleClick
        fetchIDandDes()

    End Sub



    Sub closeThisForm()
        recvGoodsMain.Enabled = True
        Home_Page.Enabled = True

        With reordersetup
            .dgorder.ClearSelection()
            .dgorder.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub T1_TextChanged(sender As Object, e As EventArgs) Handles t1.TextChanged
        loadLV()
    End Sub

    Private Sub T2_TextChanged(sender As Object, e As EventArgs) Handles t2.TextChanged
        loadLV()
    End Sub

    Private Sub T1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles t1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("Invalid Product Number")
        'End If
    End Sub
End Class