Public Class inventoryAutoBuild
    Private q As New qry
    Private table As New DataTable("Table")
    Private selectedRow As Integer = -1
    Public approvBy As String = ""

    Private Sub inventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpDate.MaxDate = DateTime.Now
        newDGV()
        q.fetchreasonnewcount(ComboBox1)
        cbxOwnership.SelectedIndex = 0
    End Sub

    Private Sub newDGV()
        'With table
        '    .Columns.Add("PRODUCT NUMBER", Type.GetType("System.String"))  '0
        '    .Columns.Add("batchId", Type.GetType("System.String"))     '1
        '    .Columns.Add("locId", Type.GetType("System.String"))  '2
        '    .Columns.Add("DESCRIPTION", Type.GetType("System.String"))     '3
        '    .Columns.Add("BATCH CODE", Type.GetType("System.String"))  '4
        '    .Columns.Add("LOCATION", Type.GetType("System.String"))     '5
        '    .Columns.Add("QUANTITY", Type.GetType("System.String"))     '6
        '    .Columns.Add("EXPIRATION DATE", Type.GetType("System.String"))     '7
        'End With
        With dgvInventory

            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35

            .Columns.Add("0", "PRODUCT NUMBER")
            .Columns.Add("1", "batchId")
            .Columns.Add("2", "locId")
            .Columns.Add("3", "DESCRIPTION")
            .Columns.Add("4", "BATCH CODE")
            .Columns.Add("5", "LOCATION")
            .Columns.Add("6", "QUANTITY")
            .Columns.Add("7", "EXPIRATION DATE")

            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = True
            .EnableHeadersVisualStyles = False
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35

            .Columns(0).Width = 125
            .Columns(1).Visible = False
            .Columns(2).Visible = False
            .Columns(3).Width = 570
            .Columns(4).Width = 90
            .Columns(5).Width = 90
            .Columns(6).Width = 96
            .Columns(7).Width = 170
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = False
            .Columns(7).ReadOnly = True

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
                .ColumnHeadersDefaultCellStyle.Font = New Font(adjustInventory.dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
                .DefaultCellStyle.Font = New Font(adjustInventory.dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Regular)
            Next
            .ClearSelection()
        End With
    End Sub

    Private Sub inventory_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If (dgvInventory.Rows.Count > 0 And btnSave.Text = "Record") Or (Not String.IsNullOrWhiteSpace(tbxDocNum.Text) And btnSave.Text = "Record") Then
            Dim result As Integer = MessageBox.Show("Are you sure want cancel transaction?", "Remove Details", MessageBoxButtons.YesNo)
            If result = DialogResult.No Then
                e.Cancel = True
            ElseIf result = DialogResult.Yes Then
                e.Cancel = False
                dgvInventory.Rows.Clear()
                dgvInventory.Columns.Clear()
            End If
        End If



    End Sub

    Private Sub dgvInventory_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInventory.CellDoubleClick
        Dim row As Integer = dgvInventory.CurrentCell.RowIndex
        If e.ColumnIndex = 0 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectProdABProducts
                .Show()
                .targetRow = row
            End With

        ElseIf e.ColumnIndex = 4 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectProdABBatch
                .Show()
                .targetRow = row
            End With

        ElseIf e.ColumnIndex = 5 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectProdABLocations
                .Show()
                .targetRow = row
            End With
        ElseIf e.ColumnIndex = 7 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectExpirationdate
                .Show()
                .rowToEdit = row
            End With
        End If
    End Sub

    'Private Sub dgvInventory_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvInventory.CellValidating
    '    If e.ColumnIndex = 6 Then
    '        With dgvInventory
    '            If .IsCurrentCellDirty Then
    '                If Not IsNumeric(e.FormattedValue) Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                ElseIf e.FormattedValue < 1 Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End If
    '            End If
    '        End With
    '    End If
    'End Sub

    Private Sub dgvInventory_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvInventory.RowValidating
        Dim row As DataGridViewRow = dgvInventory.Rows(e.RowIndex)
        Dim isNull As Boolean = False
        Dim repeatingId As Boolean
        With dgvInventory
            Try
                If .IsCurrentRowDirty Then
                    For Each cell As DataGridViewCell In row.Cells
                        If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                            e.Cancel = True
                            isNull = True
                        End If
                    Next

                    For i As Integer = 0 To Me.dgvInventory.RowCount - 1
                        For j As Integer = 0 To Me.dgvInventory.RowCount - 1
                            If i <> j Then
                                If String.Concat(dgvInventory.Rows(i).Cells(0).Value, dgvInventory.Rows(i).Cells(1).Value, dgvInventory.Rows(i).Cells(2).Value) = String.Concat(dgvInventory.Rows(j).Cells(0).Value, dgvInventory.Rows(j).Cells(1).Value, dgvInventory.Rows(j).Cells(2).Value) Then
                                    dgvInventory.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                    e.Cancel = True
                                    lblErr.Visible = True
                                    lblErr.Text = "Duplicate data are not valid."
                                    repeatingId = True
                                Else
                                    dgvInventory.Rows(i).DefaultCellStyle.BackColor = Color.White
                                End If
                            End If
                        Next
                    Next

                    If isNull = True Then
                        MessageBox.Show("Fields with * are required.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf repeatingId = False Then
                        lblErr.Visible = False
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End With
    End Sub

    Sub validateDatasToAdd()
        If dgvInventory.Rows.Count > 1 And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Or String.IsNullOrWhiteSpace(ComboBox1.Text) Then
            Home_Page.Enabled = False
            Me.Enabled = False
            inventoryAutoBuildConfirmation.Show()
        Else
            'MsgBox("No Datas")
            MessageBox.Show("Fields with * are required.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Sub addDatas()
        q.addProductAutoBuild(tbxDocNum.Text, dtpDate.Value, ComboBox1.Text, cbxOwnership.Text, newHome.userId, newHome.areaId, approvBy)
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        If btnSave.Text = "Record" Then

            If tbxDocNum.Text = "" Then
                MessageBox.Show("Please Enter Document Number...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf ComboBox1.Text = "" Then
                MessageBox.Show("Please Put Remarks...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else

                If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    validateDatasToAdd()
                End If

            End If

        ElseIf btnSave.Text = "Add New Entry" Then
                With Me
                .dgvInventory.AllowUserToAddRows = True
                .dgvInventory.Enabled = True
                .table.Clear()
                .lblErr.Visible = False
                .lblErr.ForeColor = Color.Red
                .btnSave.Text = "Record"
                .tbxEntryNumber.Text = ""
                .dtpDate.Value = Date.Now
                .ComboBox1.SelectedIndex = -1
                .cbxOwnership.SelectedIndex = 0
                .btnSave.Select()
                .tbxDocNum.Text = ""
                ComboBox1.SelectedItem = -1
            End With
        End If
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub tbxDocNum_Leave(sender As Object, e As EventArgs) Handles tbxDocNum.Leave
        q.validateInventoryAutoBuildDocNum(tbxDocNum.Text, newHome.areaId)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub dgvInventory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvInventory.CellContentClick

    End Sub

    Private Sub Expirationdate_ValueChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub InventoryAutoBuild_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class