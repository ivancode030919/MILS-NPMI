Public Class selectExpirationdate



    Private columnLVSelected As Integer = 1
    Public rowToEdit As Integer
    Private locName As String
    Private locId As String
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            DateTimePicker1.Visible = False
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        fetchIDandDes()
    End Sub


    Sub fetchIDandDes()

        If RadioButton1.Checked = True Then
            locId = TextBox1.Text
            With inventoryAutoBuild
                .dgvInventory.Rows(rowToEdit).Cells(7).Value = locId
                closeThisForm()
                .dgvInventory.ClearSelection()
            End With

        ElseIf RadioButton2.Checked = True Then

            locName = DateTimePicker1.Value.ToString("MM/dd/yyyy")
            With inventoryAutoBuild
                .dgvInventory.Rows(rowToEdit).Cells(7).Value = locName
                closeThisForm()
                .dgvInventory.ClearSelection()
            End With
        End If
    End Sub



    Sub closeThisForm()
        inventoryAutoBuild.Enabled = True
        Home_Page.Enabled = True
        With inventoryAutoBuild
            .dgvInventory.ClearSelection()
            .dgvInventory.Select()
        End With
        Me.Dispose()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        closeThisForm()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            DateTimePicker1.Visible = True
        End If
    End Sub

    Private Sub SelectExpiration_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker1.Value = Date.Now.ToShortDateString
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MM/dd/yyyy"
    End Sub

    Private Sub SelectExpiration_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        closeThisForm()
    End Sub
End Class