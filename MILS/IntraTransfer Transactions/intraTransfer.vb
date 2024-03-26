Public Class intraTransfer
    Private q As New qry
    Private n As New qryv3
    Private table As New DataTable("Table")
    Dim selectedRow As Integer = -1
    Public approvedId As String = ""
    Public userxxx As String = ""
    Private slctedRow As Integer
    Public qty As Double
    Public dgvQty As Double

    Private Sub intraTransfer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpTransferDate.MaxDate = DateTime.Now
        loadUser()
        newDGVFormat()
        q.fetchreasontran(ComboBox1)
    End Sub

    Private Sub tbxTransferId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxTransferId.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Sub loadUser()
        q.fetchUserLogIn(newHome.userId)
    End Sub

    Sub newDGVFormat()

        With dgvTransferGoods
            .Columns.Clear()
            '.AllowUserToResizeColumns = False
            '.AllowUserToResizeRows = False
            '.AllowUserToAddRows = True
            '.EnableHeadersVisualStyles = False
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.DataSource = table

            .Columns.Add("0", "pId")
            .Columns.Add("1", "bId")
            .Columns.Add("2", "locIdF")
            .Columns.Add("3", "locIdT")
            .Columns.Add("4", "aId")
            .Columns.Add("5", "PRODUCT NUMBER")
            .Columns.Add("6", "DESCRIPTION")
            .Columns.Add("7", "BATCH CODE")
            .Columns.Add("8", "FROM")
            .Columns.Add("9", "TO")
            .Columns.Add("10", "QUANTITY")
            .Columns.Add("11", "EXPIRATION DATE")
            .Columns.Add("12", "QTY")

            .Columns(0).Visible = False
            .Columns(1).Visible = False
            .Columns(2).Visible = False
            .Columns(3).Visible = False
            .Columns(4).Visible = False
            .Columns(12).Visible = False
            .Columns(0).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 523
            .Columns(7).Width = 90
            .Columns(8).Width = 90
            .Columns(10).Width = 90
            .Columns(11).Width = 150
            .Columns(12).Width = 150
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = True
            .Columns(7).ReadOnly = True
            .Columns(8).ReadOnly = True
            .Columns(9).ReadOnly = True
            .Columns(10).ReadOnly = False
            .Columns(11).ReadOnly = True

            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9, FontStyle.Bold)
                .Columns(Index).DefaultCellStyle.Font = New Font(.DefaultCellStyle.Font.FontFamily, 9)
                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .ClearSelection()
        End With
    End Sub

    Private Sub dgvTransferGoods_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTransferGoods.CellDoubleClick

        Dim rows As DataGridViewRow = dgvTransferGoods.Rows(e.RowIndex)
            Dim row As Integer = dgvTransferGoods.CurrentCell.RowIndex
        'Dim locID As String = rows.Cells(2).Value.ToString

        If e.ColumnIndex = 5 Then
            With intrTransferLoadProducts
                .targetRow = row
                .loadLV()
                .tbxDescrip.Text = ""
                .tbxProdId.Text = ""
                .tbxProdId.Select()
                .Show()

                Me.Enabled = False
            End With
        ElseIf e.ColumnIndex = 9 Then
            Dim locID As String = rows.Cells(2).Value.ToString

            If String.IsNullOrWhiteSpace(rows.Cells(0).Value.ToString) Then
                    MessageBox.Show("Select valid product first.", "Error",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    With intraTransferLoadLocationTo
                        .targetRow = row
                        .locId = locID
                        .loadLV()
                        .tbxFilter.Text = ""
                        .tbxFilter.Select()
                        .Show()

                        Me.Enabled = False
                    End With
                End If
            End If


    End Sub

    Private Sub dgvTransferGoods_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvTransferGoods.RowValidating
        Dim row As DataGridViewRow = dgvTransferGoods.Rows(e.RowIndex)
        Dim isValidQty As Boolean = True
        Dim isEmpty As Boolean = False
        Dim isNotRepeatingId As Boolean = False

        Try
            With dgvTransferGoods
                If .IsCurrentRowDirty Then
                    Dim idDGV As String = row.Cells(0).Value.ToString
                    Dim qtyDGV As String = row.Cells(10).Value.ToString
                    For Each cell As DataGridViewCell In row.Cells 'is null
                        If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                            e.Cancel = True
                            isEmpty = True
                        End If
                    Next

                    'For i As Integer = 0 To Me.dgvTransferGoods.RowCount - 1
                    '    For j As Integer = 0 To Me.dgvTransferGoods.RowCount - 1
                    '        If i <> j Then
                    '            If dgvTransferGoods.Rows(i).Cells(0).Value = dgvTransferGoods.Rows(j).Cells(0).Value Then
                    '                dgvTransferGoods.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    '                e.Cancel = True
                    '                'lblErr.Visible = True
                    '                'lblErr.Text = "Duplicate data are not valid."
                    '                isNotRepeatingId = True
                    '            Else
                    '                dgvTransferGoods.Rows(i).DefaultCellStyle.BackColor = Color.White
                    '            End If
                    '        End If
                    '    Next
                    'Next

                    Try
                        isValidQty = q.validationQuantityProd(idDGV, qtyDGV) ' check the quantity
                    Catch ex As Exception
                        MessageBox.Show("Fields with * are required.", "Error",
                                           MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try

                    If isValidQty = False Then
                        e.Cancel = True
                        MessageBox.Show("You do not have enough of these items.", "Error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf isEmpty = True Then
                        MessageBox.Show("Fields with * are required.", "Error",
                                               MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf isNotRepeatingId = False Then
                        'lblErr.Visible = False
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Private Sub dgvTransferGoods_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvTransferGoods.CellValidating
    '    Dim row As DataGridViewRow = dgvTransferGoods.Rows(e.RowIndex)
    '    If e.ColumnIndex = 10 Then
    '        With dgvTransferGoods
    '            If .IsCurrentCellDirty Then
    '                If Not IsNumeric(e.FormattedValue) Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                ElseIf e.FormattedValue = 0 Or e.FormattedValue < 0 Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

    '                End If

    '            End If
    '        End With
    '    End If
    'End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If btnAdd.Text = "Record" Then
            If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                validateFields()
            End If
        ElseIf btnAdd.Text = "Add New Entry" Then
            With Me
                .dgvTransferGoods.Enabled = True
                .dgvTransferGoods.AllowUserToAddRows = True
                .btnAdd.Text = "Record"
                .clearFields()
                .tbxTransferId.Select()
            End With
            With intrTransferLoadProducts
                .loadLV()
            End With
        End If
    End Sub

    Sub clearFields()
        table.Clear()
        tbxEntryNumber.Text = ""
        dtpTransferDate.Value = Date.Now
        ComboBox1.SelectedItem = -1
        tbxTransferId.Text = ""
        'tbxUser.Text = ""
        dgvTransferGoods.Rows.Clear()
        loadUser()
    End Sub

    Sub validateFields()
        Dim dgvCount As Integer = dgvTransferGoods.Rows.Count - 1
        If String.IsNullOrWhiteSpace(tbxTransferId.Text) Or dgvCount = 0 Then
            'lblErr.Visible = True
            'lblErr.ForeColor = Color.Red
            'lblErr.Text = "Fields with * are required."
            MessageBox.Show("Fields with * are required !!!")
            Exit Sub
        Else
            'lblErr.Visible = False
            With intraTransferConfirmation

                Me.Enabled = False
                .Show()
            End With
            '
        End If
    End Sub

    Sub addDatas()
        q.addIntraTransfer(tbxTransferId.Text, dtpTransferDate.Value, userxxx, ComboBox1.Text, newHome.userId, newHome.areaId, approvedId)
    End Sub
    Private Sub dgvTransferGoods_Click(sender As Object, e As EventArgs) Handles dgvTransferGoods.Click
        Try
            If dgvTransferGoods.Rows.Count = 0 Then
                'lblErr.Visible = True
                'lblErr.ForeColor = Color.Red
                'lblErr.Text = "Please select valid data..."
                MessageBox.Show("Please select valid data...")
            Else
                'lblErr.Visible = False
                selectedRow = dgvTransferGoods.CurrentCell.RowIndex
            End If
        Catch ex As Exception
            MessageBox.Show("Please select valid data...")
        End Try
    End Sub

    Private Sub intraTransfer_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If dgvTransferGoods.Rows.Count > 0 Then
            If e.KeyCode = Keys.Delete Then
                Dim result As Integer = MessageBox.Show("Are you sure want to remove selected row?", "Remove Details", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    'Exit Sub
                ElseIf result = DialogResult.Yes Then
                    Try
                        table.Rows.RemoveAt(selectedRow)
                    Catch ex As Exception

                    End Try
                End If
            End If

            If e.KeyCode = Keys.Escape Then
                Me.Close()
            End If

        Else

            If e.KeyCode = Keys.Escape Then
                Me.Close()
            End If

        End If
    End Sub

    Private Sub intraTransfer_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If btnAdd.Text = "Add New Entry" Then
            e.Cancel = False

        Else
            If Not String.IsNullOrWhiteSpace(tbxTransferId.Text) Or dgvTransferGoods.Rows.Count > 0 And dgvTransferGoods.Enabled = True Then
                If (MessageBox.Show("Are you sure you want to cancel this transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    e.Cancel = True
                Else
                    e.Cancel = False


                    dgvTransferGoods.Rows.Clear()
                End If
            Else

            End If
            Home_Page.MenuStrip1.Enabled = True
        End If

    End Sub

    Private Sub tbxTransferId_Leave(sender As Object, e As EventArgs) Handles tbxTransferId.Leave
        If Not String.IsNullOrWhiteSpace(tbxTransferId.Text) Then
            q.validateDocumentNumberIntraTransfer(tbxTransferId.Text, newHome.areaId)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    'checking for duplicate pd in rows
    Public Sub PDchecking()
        Dim repeatedValues As New List(Of Integer)
        Dim sumByFcount As New Dictionary(Of Integer, Integer)()

        For i As Integer = 0 To Me.dgvTransferGoods.RowCount - 1
            Dim fcount As Integer = CInt(dgvTransferGoods.Rows(i).Cells(0).Value)
            Dim valueToAdd As Integer = CInt(dgvTransferGoods.Rows(i).Cells(10).Value)

            ' Check if fcount is already in the list
            If repeatedValues.Contains(fcount) Then
                ' Update the sum for this fcount
                sumByFcount(fcount) += valueToAdd
                n.Transferid(fcount)
                ' Update the value in the cell
                dgvTransferGoods.Rows(i).Cells(12).Value = sumByFcount(fcount)
                dgvQty = sumByFcount(fcount)
            Else
                ' Add fcount to the list and initialize its sum
                repeatedValues.Add(fcount)
                sumByFcount.Add(fcount, valueToAdd)
            End If
        Next
    End Sub




    Private Sub dgvTransferGoods_RowValidated(sender As Object, e As DataGridViewCellEventArgs) Handles dgvTransferGoods.RowValidated

        PDchecking()

        If dgvQty > qty Then
            MessageBox.Show("You do not have enough of these items.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            dgvTransferGoods.AllowUserToAddRows = False
            btnAdd.Enabled = False
        ElseIf dgvQty = qty Or dgvQty < qty Then
            dgvTransferGoods.AllowUserToAddRows = True
            btnAdd.Enabled = True
        End If

    End Sub
End Class