Public Class NewCountSelected

    Private columnLVSelected As Integer = 0
    Private q As New qry
    Public AddRow As Integer
    Private pdno As String
    Private prodDes As String
    Private batch As String
    Private loc As String
    Private onhand As String
    Private counted As String
    Private diff As String


    Private Sub NewCountSelected_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVformat()
        newLVformat2()
        loaddata()
    End Sub

    Sub loaddata()
        q.loadcountdata(ListView1, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox4.Text, columnLVSelected)
    End Sub


    Sub newLVformat()
        With ListView1
            .Columns.Add("id#", 0)
            .Columns.Add("PRODUCT NUMBER", 120)
            .Columns.Add("DESCRIPTION", 450)
            .Columns.Add("BATCH", 70)
            .Columns.Add("LOCATION", 80)
            .Columns.Add("AREA", 120)
            .Columns.Add("EXPIRATION DATE", 120)
            .Columns.Add("ON HAND", 0)
            .Columns.Add("COUNTED", 0)
            .Columns.Add("DIFFERENCE", 0)
        End With
    End Sub

    Sub newLVformat2()
        With ListView2
            .Columns.Add("id#", 0)
            .Columns.Add("PRODUCT NUMBER", 120)
            .Columns.Add("DESCRIPTION", 450)
            .Columns.Add("BATCH", 70)
            .Columns.Add("LOCATION", 80)
            .Columns.Add("EXPIRATION DATE", 120)
            .Columns.Add("ON HAND", 0)
            .Columns.Add("COUNTED", 0)
            .Columns.Add("DIFFERENCE", 0)

        End With
    End Sub
    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        For Each selectedItem As ListViewItem In ListView1.SelectedItems
            ' Clone the ListViewItem
            Dim clonedItem As ListViewItem = DirectCast(selectedItem.Clone(), ListViewItem)

            ' Exclude the subitem at index 4 (column 5)
            clonedItem.SubItems.RemoveAt(5)

            ' Add the cloned item to ListView2
            ListView2.Items.Add(clonedItem)
        Next
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loaddata()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        loaddata()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        loaddata()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        loaddata()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        For Each selectedItem As ListViewItem In ListView1.Items
            ListView2.Items.Add(selectedItem.Clone())
        Next
    End Sub

    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        If ListView2.SelectedItems.Count > 0 Then
            Dim item As ListViewItem = ListView2.SelectedItems(0)
            ListView2.Items.Remove(item)
        End If
    End Sub

    Private Sub ListView2_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView2.ColumnClick
        If ListView2.SelectedItems.Count > 0 Then
            Dim item As ListViewItem = ListView2.SelectedItems(0)
            ListView2.Items.Remove(item)
        End If
    End Sub

    Private Sub ListView2_Click(sender As Object, e As EventArgs) Handles ListView2.Click

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ListView2.Items.Clear()
    End Sub

    Sub fetchDatas()

        'Dim form2 As New adjustInventory()
        'form2.newDGVFormat() ' Call the formatting method in Form2

        For Each item As ListViewItem In ListView2.Items
            Dim rowData(8) As String

            For i As Integer = 0 To 8
                rowData(i) = item.SubItems(i).Text
            Next

            adjustInventory.AddRow(rowData)
        Next

        adjustInventory.Show()
        Me.Close()
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fetchDatas()
        'adjustInventory.Show()
        'With adjustInventory.dgvProducts
        '    .Columns("DESCRIPTION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'End With

        Me.Dispose()
        adjustInventory.Enabled = True

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Dispose()
        adjustInventory.Show()
    End Sub

    Private Sub NewCountSelected_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.Dispose()
        adjustInventory.Show()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class