Public Class reserveRelease
    Private q As New qry
    Private table As New DataTable("Table")
    Public docTypeId As String
    Public docRefTypeId As String
    Public receiverId As String
    Private slctedRow As Integer

    Private Sub reserveRelease_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newDGVFormat()
        loadForm()
    End Sub

    Sub loadForm()
        loadTbxDocType()
        loadtbxRefDocType()
        q.loadSender(cbxReceiver)
        cbxOwnership.SelectedIndex = 0
    End Sub

    Sub newDGVFormat()
        With table
            .Columns.Add("id", Type.GetType("System.String"))  '0
            .Columns.Add("PRODUCT NUMBER", Type.GetType("System.String"))  '1
            .Columns.Add("DESCRIPTION", Type.GetType("System.String"))     '2
            .Columns.Add("BATCH CODE", Type.GetType("System.String"))  '3
            .Columns.Add("LOCATION", Type.GetType("System.String"))     '4
            .Columns.Add("QUANTITY", Type.GetType("System.String"))     '5
            .Columns.Add("bId", Type.GetType("System.String"))     '6
            .Columns.Add("lId", Type.GetType("System.String"))     '7
        End With
        With dgvRels
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = True
            .EnableHeadersVisualStyles = False
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            .DataSource = table
            .Columns(0).Visible = False
            .Columns(1).Width = 150
            .Columns(2).Width = 650
            .Columns(3).Width = 125
            .Columns(4).Width = 125
            .Columns(5).Width = 152
            .Columns(6).Visible = False
            .Columns(7).Visible = False
            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = False
            .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .ClearSelection()
        End With
    End Sub

    Sub loadTbxDocType()
        With tbxDocType
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            'q.suggestDocType(col)
            q.suggestDocTypeRels(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Sub loadtbxRefDocType()
        With tbxRefDocType
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            'q.suggestRefDocType(col)
            q.suggestRefDocTypeRels(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Private Sub cbxReceiver_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxReceiver.SelectedIndexChanged
        receiverId = ""
        receiverId = q.fetchSenderId(cbxReceiver.Text)
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbxRefDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxRefDocNum.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Sub ValidateDocumentNumber()
        q.validateDocumentReservRelsTransaction(tbxDocType.Text, docTypeId, tbxDocNum.Text, newHome.areaId)
    End Sub

    Private Sub tbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) Then
            'q.fetchIdDocType(tbxDocType.Text)
            q.fetchIdDocTypeReservRels(tbxDocType.Text)
            'MsgBox(docTypeId)
        Else

        End If

        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Then
            ValidateDocumentNumber()
        End If
    End Sub

    Private Sub tbxRefDocType_Leave(sender As Object, e As EventArgs) Handles tbxRefDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) Then
            'q.fetchIdDocType(tbxDocType.Text)
            q.fetchIdRefDocTypeReserveRels(tbxRefDocType.Text)
            'MsgBox(docRefTypeId)
        Else

        End If
    End Sub

    'Private Sub dgvRels_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvRels.CellValidating
    '    Dim row As DataGridViewRow = dgvRels.Rows(e.RowIndex)
    '    If e.ColumnIndex = 5 Then
    '        With dgvRels
    '            If .IsCurrentCellDirty Then
    '                If Not IsNumeric(e.FormattedValue) Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                ElseIf e.FormattedValue = 0 Or e.FormattedValue < 0 Then
    '                    e.Cancel = True
    '                    'row.Cells(6).Value = emptyCell
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End If
    '            End If
    '        End With
    '    End If
    'End Sub

    Private Sub dgvRels_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRels.CellDoubleClick
        Dim row As Integer = dgvRels.CurrentCell.RowIndex
        If e.ColumnIndex = 1 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectProdforReserveRelease
                .loadLV()
                .Show()
                .targetRow = row
            End With
        End If
    End Sub

    Private Sub dgvRels_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvRels.RowValidating
        Dim row As DataGridViewRow = dgvRels.Rows(e.RowIndex)
        Dim isValidQty As Boolean = True
        Dim repeatingId As Boolean = False
        Dim isEmpty As Boolean

        Try
            With dgvRels
                If .IsCurrentRowDirty Then
                    Dim idDGV As String = ""
                    idDGV = row.Cells(0).Value.ToString
                    Dim qtyDGV As String = ""
                    qtyDGV = row.Cells(5).Value.ToString

                    For Each cell As DataGridViewCell In row.Cells 'is null
                        If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                            e.Cancel = True
                            isEmpty = True
                        End If
                    Next

                    For i As Integer = 0 To Me.dgvRels.RowCount - 1
                        For j As Integer = 0 To Me.dgvRels.RowCount - 1
                            If i <> j Then
                                If dgvRels.Rows(i).Cells(0).Value = dgvRels.Rows(j).Cells(0).Value Then
                                    dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                    e.Cancel = True
                                    lblErr.Visible = True
                                    lblErr.Text = "Duplicate data are not valid."
                                    repeatingId = True
                                Else
                                    dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.White
                                End If
                            End If
                        Next
                    Next

                    Try
                        isValidQty = q.validationQuantityProd(idDGV, qtyDGV) ' check the quantity
                    Catch ex As Exception
                        MessageBox.Show("Fields with * are required.", "Error",
                                           MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                    If isEmpty = True Then
                        e.Cancel = True
                        MessageBox.Show("Fields with * are required.", "Error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf isValidQty = False Then
                        e.Cancel = True
                        MessageBox.Show("You do not have enough of these items.", "Error",
                                                MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf repeatingId = False Then
                        lblErr.Visible = False
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub dgvRels_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvRels.KeyDown
        If dgvRels.Rows.Count > 1 Then
            If e.KeyCode = Keys.Delete Then
                Dim result As Integer = MessageBox.Show("Are you sure want to remove selected row?", "Remove Details", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    'Exit Sub
                ElseIf result = DialogResult.Yes Then
                    Try
                        table.Rows.RemoveAt(slctedRow)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub dgvRels_Click(sender As Object, e As EventArgs) Handles dgvRels.Click
        If dgvRels.Rows.Count = 0 Then
            lblErr.Visible = True
            lblErr.ForeColor = Color.Red
            lblErr.Text = "Please select valid data..."
        Else
            If dgvRels.Rows.Count > 1 Then
                lblErr.Visible = False
                slctedRow = dgvRels.CurrentCell.RowIndex
            End If
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If btnAdd.Text = "Record" Then
            validateFields()
        ElseIf btnAdd.Text = "Add New Entry" Then
            With Me
                .dgvRels.AllowUserToAddRows = True
                .dgvRels.Enabled = True
            End With
            clearfields()
            cbxReceiver.Select()
            btnAdd.Text = "Record"
            lblErr.Visible = False
        End If
    End Sub

    Sub clearfields()
        cbxReceiver.SelectedIndex = -1
        tbxDocType.Text = ""
        docTypeId = ""
        tbxDocNum.Text = ""
        tbxRemarks.Text = ""
        tbxEntry.Text = ""
        tbxRefDocType.Text = ""
        docRefTypeId = ""
        tbxRefDocNum.Text = ""
        cbxOwnership.SelectedIndex = 0
        table.Clear()
        dtpRefDate.Value = Date.Now
    End Sub

    Sub validateFields()
        Dim dgvCount As Integer = dgvRels.Rows.Count - 1
        If cbxReceiver.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or String.IsNullOrWhiteSpace(tbxDocNum.Text) Or String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or dgvCount = 0 Then

            lblErr.Visible = True
            lblErr.ForeColor = Color.Red
            lblErr.Text = "Fields with * are required."
            Exit Sub
        Else
            lblErr.Visible = False
            q.addReserveRelsTransaction(docTypeId, tbxDocNum.Text, dtpRefDate.Value, newHome.userId, receiverId,
                                 tbxRemarks.Text, cbxOwnership.Text, docRefTypeId, tbxRefDocNum.Text, newHome.areaId)
        End If

    End Sub

    Private Sub reserveRelease_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If btnAdd.Text = "Add New Entry" Then
            e.Cancel = False
        Else
            If Not cbxReceiver.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Or Not String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or dgvRels.Rows.Count > 1 Then
                If (MessageBox.Show("Are you sure you want to cancel this transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    e.Cancel = True
                Else
                    e.Cancel = False
                End If
            Else
                e.Cancel = False
            End If
        End If
    End Sub

    Private Sub tbxDocNum_Leave(sender As Object, e As EventArgs) Handles tbxDocNum.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Then
            ValidateDocumentNumber()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class