Public Class adjustInventory
    Private q As New qry
    Private dt As New DataTable
    Private columnLVSelected As Integer = 0
    Private isEditedDGV As Boolean = False
    Public approvedID As String = ""

    Private Sub adjustInventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        isEditedDGV = False
        dtpDate.MaxDate = DateTime.Now

        RadioButton1.Checked = True
        q.fetchUserLogInAdjustInventory(newHome.userId)
        q.fetchrAreaAdjustInventory(newHome.areaId)
        q.fetchreasoninadjustment(ComboBox1)
    End Sub


    Public Sub AddRow(rowData() As String)
        dt.Rows.Add(rowData)
        dgvProducts.Columns(0).Visible = False
    End Sub

    Private Sub loadDGV()

        q.loadDGVInventoryAdjust(TextBox1.Text, TextBox2, TextBox3, newHome.areaId)
        afterDataBind()
    End Sub

    Sub afterDataBind()
        If dgvProducts.Rows.Count < 1 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
            btnSave.Enabled = False
        Else
            lblErr.Visible = False
            btnSave.Enabled = True
        End If
    End Sub

    Sub validateActionforAdding()
        Dim editedDGV = False
        With dgvProducts
            For i As Integer = 0 To .Rows.Count - 1 Step +1
                If .Rows(i).Cells(8).Value <> 0 Then
                    editedDGV = True
                End If
            Next
        End With

        If editedDGV = True Then
            Home_Page.Enabled = False
            Me.Enabled = False
            confirmationFormAdjustInventory.Show()
        Else
            MessageBox.Show("Please modify some data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


    End Sub

    'Private Sub dgvProducts_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs)
    '    Dim row As DataGridViewRow = dgvProducts.Rows(e.RowIndex)
    '    Dim countedQty As Double = 0
    '    Dim onhandQty As Double = 0
    '    Dim differenceQty As Double = 0
    '    Try
    '        If e.ColumnIndex = 7 Then
    '            With dgvProducts

    '                countedQty = row.Cells(7).Value.ToString
    '                onhandQty = e.FormattedValue
    '                differenceQty = onhandQty - countedQty

    '                row.Cells(8).Value = differenceQty.ToString
    '                'End If
    '            End With
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show("Please Input Counted Quantity!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    'Private Sub dgvProducts_Click(sender As Object, e As EventArgs)
    '    isEditedDGV = True
    'End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If btnSave.Text = "Record" Then
            If tbxDocnum.Text = "" Then
                MessageBox.Show("Please Enter Document Number...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf ComboBox1.Text = "" Then
                MessageBox.Show("Please Put Remarks...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    validateActionforAdding()
                End If

            End If
        ElseIf btnSave.Text = "Add New" Then
            isEditedDGV = True
            dgvProducts.Enabled = True
            tbxEntryNumber.Text = ""
            ComboBox1.SelectedIndex = -1
            tbxDocnum.Text = ""
            q.fetchUserLogInAdjustInventory(newHome.userId)
            dtpDate.Value = Date.Now
            btnSave.Text = "Record"
            Me.Refresh()
            RadioButton1.Checked = True
            TextBox1.Text = "1"
            TextBox1.Text = ""
        End If
    End Sub

    Sub addDatas()
        q.addAdjustInventory(tbxDocnum.Text, ComboBox1.Text, tbxUser.Text, dtpDate.Value, approvedID, newHome.userId, newHome.areaId)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
        isEditedDGV = True
        dgvProducts.Enabled = True
        tbxEntryNumber.Text = ""
        ComboBox1.SelectedIndex = -1
        tbxDocnum.Text = ""
        q.fetchUserLogInAdjustInventory(newHome.userId)
        dtpDate.Value = Date.Now
        btnSave.Text = "Record"
        Me.Refresh()
        RadioButton1.Checked = True
        TextBox1.Text = "1"
        TextBox1.Text = ""
    End Sub

    'Private Sub adjustInventory_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
    '    Dim editedDGV = False
    '    With dgvProducts
    '        For i As Integer = 0 To .Rows.Count - 1 Step +1
    '            If .Rows(i).Cells(8).Value <> 0 Then
    '                editedDGV = True
    '            End If
    '        Next
    '    End With

    '    If editedDGV = True And btnSave.Text = "Record" Then
    '        If (MessageBox.Show("Are you sure to cancel transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
    '            e.Cancel = True
    '        Else
    '            e.Cancel = False
    '        End If
    '    Else
    '        e.Cancel = False
    '    End If

    'End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim NewCountSelected As New NewCountSelected()
        NewCountSelected.ShowDialog()

    End Sub




    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        If dt.Columns.Contains("ID") Then
            dt.Columns.Remove("ID")
        End If

        If dt.Columns.Contains("PRODUCT NUMBER") Then
            dt.Columns.Remove("PRODUCT NUMBER")
        End If

        If dt.Columns.Contains("DESCRIPTION") Then
            dt.Columns.Remove("DESCRIPTION")
        End If


        If dt.Columns.Contains("BATCH") Then
            dt.Columns.Remove("BATCH")
        End If


        If dt.Columns.Contains("LOCATION") Then
            dt.Columns.Remove("LOCATION")
        End If

        If dt.Columns.Contains("EXPIRATION DATE") Then
            dt.Columns.Remove("EXPIRATION DATE")
        End If


        If dt.Columns.Contains("ON HAND") Then
            dt.Columns.Remove("ON HAND")
        End If


        If dt.Columns.Contains("COUNTED") Then
            dt.Columns.Remove("COUNTED")
        End If


        If dt.Columns.Contains("DIFFERENCE") Then
            dt.Columns.Remove("DIFFERENCE")
        End If

        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            Button1.Visible = False
            loadDGV()

        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            Button1.Visible = True
            dgvProducts.DataSource = Nothing
            newDGVFormat()
        End If
    End Sub

    Public Sub newDGVFormat()
        ' Column names and types
        dt.Rows.Clear()
        dt.Columns.Add("ID", GetType(String))
        dt.Columns.Add("PRODUCT NUMBER", GetType(String))
        dt.Columns.Add("DESCRIPTION", GetType(String))
        dt.Columns.Add("BATCH", GetType(String))
        dt.Columns.Add("LOCATION", GetType(String))
        dt.Columns.Add("EXPIRATION DATE", GetType(String))
        dt.Columns.Add("ON HAND", GetType(String))
        dt.Columns.Add("COUNTED", GetType(String))
        dt.Columns.Add("DIFFERENCE", GetType(String))

        dgvProducts.DataSource = dt

        'Set column widths And other formatting settings
        dgvProducts.Columns("ID").Width = 0
        dgvProducts.Columns("PRODUCT NUMBER").Width = 100
        dgvProducts.Columns("DESCRIPTION").Width = 480
        dgvProducts.Columns("BATCH").Width = 100
        dgvProducts.Columns("LOCATION").Width = 100
        dgvProducts.Columns("EXPIRATION DATE").Width = 110
        dgvProducts.Columns("ON HAND").Width = 90
        dgvProducts.Columns("COUNTED").Width = 90
        dgvProducts.Columns("DIFFERENCE").Width = 110

        dgvProducts.Columns("ID").ReadOnly = False
        dgvProducts.Columns("PRODUCT NUMBER").ReadOnly = True
        dgvProducts.Columns("DESCRIPTION").ReadOnly = True
        dgvProducts.Columns("BATCH").ReadOnly = True
        dgvProducts.Columns("LOCATION").ReadOnly = True
        dgvProducts.Columns("EXPIRATION DATE").ReadOnly = True
        dgvProducts.Columns("ON HAND").ReadOnly = True
        dgvProducts.Columns("COUNTED").ReadOnly = False
        dgvProducts.Columns("DIFFERENCE").ReadOnly = True
        dgvProducts.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgvProducts.Columns("DESCRIPTION").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        'dgvProducts.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
        dgvProducts.Columns(0).Visible = False
        'dgvProducts.Columns("DESCRIPTION").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        'dgvProducts.Columns("PRODUCT NUMBER").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        For Each column As DataGridViewColumn In dgvProducts.Columns
            column.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            column.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            'dgvProducts.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            'dgvProducts.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft
            dgvProducts.Columns(0).Visible = False
        Next
        dgvProducts.ClearSelection()
    End Sub

    Private Sub TbxDocnum_TextChanged(sender As Object, e As EventArgs) Handles tbxDocnum.TextChanged
        Label9.Visible = False
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loadDGV()
    End Sub

    Private Sub TextBox2_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        loadDGV()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        loadDGV()
    End Sub

    Private Sub DgvProducts_CellValidating_1(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvProducts.CellValidating
        Dim row As DataGridViewRow = dgvProducts.Rows(e.RowIndex)
        Dim countedQty As Double = 0
        Dim onhandQty As Double = 0
        Dim differenceQty As Double = 0
        Try
            If e.ColumnIndex = 7 Then
                With dgvProducts

                    countedQty = row.Cells(6).Value.ToString
                    onhandQty = e.FormattedValue
                    differenceQty = onhandQty - countedQty

                    row.Cells(8).Value = differenceQty.ToString

                End With
            End If
        Catch ex As Exception
            MessageBox.Show("Please Input Counted Quantity!!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DgvProducts_Click(sender As Object, e As EventArgs) Handles dgvProducts.Click
        isEditedDGV = True
    End Sub

    Private Sub adjustInventory_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim editedDGV = False
        With dgvProducts
            For i As Integer = 0 To .Rows.Count - 1 Step +1
                If .Rows(i).Cells(8).Value <> 0 Then
                    editedDGV = True
                End If
            Next
        End With

        If editedDGV = True And btnSave.Text = "Record" Then
            If (MessageBox.Show("Are you sure to cancel transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                e.Cancel = True
                isEditedDGV = True
                dgvProducts.Enabled = True
                tbxEntryNumber.Text = ""
                ComboBox1.SelectedIndex = -1
                tbxDocnum.Text = ""
                q.fetchUserLogInAdjustInventory(newHome.userId)
                dtpDate.Value = Date.Now
                btnSave.Text = "Record"
                Me.Refresh()
                RadioButton1.Checked = True
                TextBox1.Text = "1"
                TextBox1.Text = ""
            Else
                e.Cancel = False
            End If
        Else
            e.Cancel = False
        End If

        Home_Page.MenuStrip1.Enabled = True

    End Sub

    Private Sub TbxUser_TextChanged(sender As Object, e As EventArgs) Handles tbxUser.TextChanged

    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress

    End Sub

    Private Sub AdjustInventory_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class