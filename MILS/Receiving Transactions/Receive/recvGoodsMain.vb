Imports Microsoft.Reporting.WinForms

Public Class recvGoodsMain
    Private q As New qry
    Private y As New qryv3
    Public docTypeId As String
    Public docRefTypeId As String
    Public senderId As String
    'Private table As New DataTable("Table")
    Private slctedRow As Integer
    Public series As String = ""
    Public areacode As String = ""
    Public docCode As String = ""
    Public branch As String = ""

    Private Sub recvGoodsMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpRefDate.MaxDate = DateTime.Now
        loadForm()
        Me.Focus()
        Me.Select()

        q.loadSender(cbxSender)
        doctypeselect()
    End Sub

    Sub loadForm()
        cbxOwnership.SelectedIndex = 0
        cbxSender.SelectedIndex = -1
        loadTbxDocType()
        loadtbxRefDocType()
        newDGVFormat()
    End Sub

    Sub loadTbxDocType()
        With tbxDocType
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestDocType(col)
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
            q.suggestRefDocType(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Private Sub tbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) Then
            q.fetchIdDocType(tbxDocType.Text)
        Else

        End If

        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Then
            validateDocuNumbers()
        Else

        End If
    End Sub

    Private Sub tbxRefDocType_Leave(sender As Object, e As EventArgs) Handles tbxRefDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxRefDocType.Text) Then
            q.fetchIdRefDocType(tbxRefDocType.Text)
        Else

        End If
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If 
    End Sub 'NOTE:this declared validation also triggers the data type set on the database refering to as the docNum data type integer which is also set as primary key auto increment

    Private Sub tbxRefDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxRefDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub 'NOTE:this declared validation also triggers the data type set on the database refering to as the docNum data type integer which is also set as primary key auto increment

    Private Sub cbxSender_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSender.SelectedIndexChanged
        If cbxSender.Text = "SUPPLIER" Then
            tbxDocType.Text = "Receiving Report"
        Else
            tbxDocType.Text = ""
        End If
        senderId = q.fetchSenderId(cbxSender.Text)
        q.fetchIdDocType(tbxDocType.Text)
    End Sub

    Sub clearfields()
        cbxSender.SelectedIndex = -1
        tbxDocType.Text = ""
        docTypeId = ""
        tbxDocNum.Text = ""
        ComboBox1.SelectedItem = -1
        tbxEntry.Text = ""
        tbxRefDocType.Text = ""
        docRefTypeId = ""
        tbxRefDocNum.Text = ""
        cbxOwnership.SelectedIndex = 0
        dgvRecv.Rows.Clear()
        TextBox1.Text = String.Empty
    End Sub

    Sub newDGVFormat()

        With dgvRecv
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

            .Columns(0).Width = 150

            .Columns(1).Visible = False
            .Columns(2).Visible = False

            .Columns(3).Width = 523
            .Columns(4).Width = 95
            .Columns(5).Width = 95
            .Columns(6).Width = 100
            .Columns(7).Width = 170

            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = False
            .Columns(7).ReadOnly = True

            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            Next


        End With
    End Sub

    'Private Sub dgvRecv_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dgvRecv.CellValidating
    '    If e.ColumnIndex = 6 Then
    '        With dgvRecv
    '            If .IsCurrentCellDirty Then
    '                If Not IsNumeric(e.FormattedValue) Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                ElseIf e.FormattedValue = 0 Or e.FormattedValue < 1 Then
    '                    e.Cancel = True
    '                    MessageBox.Show("Insert valid quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '                End If
    '            End If
    '        End With
    '    End If
    'End Sub

    Private Sub dgvRecv_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvRecv.RowValidating
        Dim row As DataGridViewRow = dgvRecv.Rows(e.RowIndex)
        Dim isNull As Boolean
        Dim repeatingId As Boolean = False
        With dgvRecv
            Try
                If .IsCurrentRowDirty Then
                    For Each cell As DataGridViewCell In row.Cells
                        If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                            e.Cancel = True
                            isNull = True
                        End If
                    Next

                    For i As Integer = 0 To Me.dgvRecv.RowCount - 1
                        For j As Integer = 0 To Me.dgvRecv.RowCount - 1
                            If i <> j Then
                                If String.Concat(dgvRecv.Rows(i).Cells(0).Value, dgvRecv.Rows(i).Cells(1).Value, dgvRecv.Rows(i).Cells(2).Value) = String.Concat(dgvRecv.Rows(j).Cells(0).Value, dgvRecv.Rows(j).Cells(1).Value, dgvRecv.Rows(j).Cells(2).Value) Then
                                    'dgvRecv.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                    e.Cancel = True
                                    lblErr.Visible = True
                                    lblErr.Text = "Duplicate data are not valid."
                                    repeatingId = True
                                Else
                                    'dgvRecv.Rows(i).DefaultCellStyle.BackColor = Color.White
                                End If
                            End If
                        Next
                    Next

                    If isNull = True Then
                        MessageBox.Show("Fields with * are required.", "Error",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ElseIf repeatingId = False Then
                        lblErr.Visible = True
                    End If
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End With
    End Sub

    Private Sub dgvRecv_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRecv.CellDoubleClick
        Dim row As Integer = dgvRecv.CurrentCell.RowIndex
        If e.ColumnIndex = 0 Then
            Me.Enabled = False
            With recvSelectGood
                .loadLV()
                .rowToEdit = row
                .Show()
                .tbxFilter.Select()
                .Focus()
            End With
        ElseIf e.ColumnIndex = 4 Then

            Me.Enabled = False
            With recvSelectBatch
                .rowToEdit = row
                .Show()
                .tbxFilter.Select()
                .Focus()
            End With
        ElseIf e.ColumnIndex = 5 Then

            Me.Enabled = False
            With recvSelectLocations
                .rowToEdit = row
                .Show()
                .tbxFilter.Select()
                .Focus()
            End With

        ElseIf e.ColumnIndex = 7 Then

            Me.Enabled = False
            With selectExpiration
                .rowToEdit = row
                .Show()
                .Focus()
            End With
        End If
    End Sub

    Sub validateFieldsForAddition()
        Dim dgvCount As Integer = dgvRecv.Rows.Count - 1

        If cbxSender.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or String.IsNullOrWhiteSpace(tbxDocNum.Text) Or String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or String.IsNullOrWhiteSpace(tbxRefDocNum.Text) Or cbxOwnership.SelectedIndex = -1 _
            Or String.IsNullOrWhiteSpace(ComboBox1.Text) Or dgvCount = 0 Then
            MessageBox.Show("Fields with '*' are required", "Info:", MessageBoxButtons.OK)
            'lblErr.Visible = True
            'lblErr.ForeColor = Color.Red
            'lblErr.Text = "Fields with '*' are required"
            Exit Sub
        Else

            q.addRecvTransaction(docTypeId, tbxDocNum.Text, dtpRefDate.Value, senderId, ComboBox1.Text, cbxOwnership.Text, docRefTypeId, tbxRefDocNum.Text, newHome.userId, newHome.areaId, series, TextBox1.Text)
        End If
    End Sub

    Sub validateDocuNumbers()
        q.validateDocumentBeforAddingRecvTrans(tbxDocType.Text, docTypeId, tbxDocNum.Text, newHome.areaId)
    End Sub

    Private Sub bntAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If btnAdd.Text = "Record" Then

            If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else

                'If cbxSender.Text = "SUPPLIER" Then
                '    q.fetchrECVsUPPLIERseries()
                'Else
                '    q.fetchSeriesRecvAndDisplay()
                'End If

                validateFieldsForAddition()
            End If

        ElseIf btnAdd.Text = "Add New Entry" Then
            With Me
                .dgvRecv.AllowUserToAddRows = True
                .dgvRecv.Enabled = True
            End With
            Button1.Visible = False
            clearfields()
            cbxSender.Select()
            btnAdd.Text = "Record"

        End If
    End Sub

    Private Sub dgvRecv_Click(sender As Object, e As EventArgs) Handles dgvRecv.Click
        If dgvRecv.Rows.Count = 0 Then
            lblErr.Visible = True
            lblErr.ForeColor = Color.Red
            lblErr.Text = "Please select valid data..."
        Else

            slctedRow = dgvRecv.CurrentCell.RowIndex
        End If
    End Sub

    Private Sub dgvRecv_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvRecv.KeyDown
        If dgvRecv.Rows.Count > 1 Then
            If e.KeyCode = Keys.Delete Then
                Dim result As Integer = MessageBox.Show("Are you sure want to remove selected row?", "Remove Details", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    'Exit Sub
                ElseIf result = DialogResult.Yes Then
                    Try
                        dgvRecv.Rows.RemoveAt(slctedRow)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub recvGoodsMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If btnAdd.Text = "Add New Entry" Then
            e.Cancel = False

        Else
            If Not cbxSender.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Or Not String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or dgvRecv.Rows.Count > 0 Then
                If (MessageBox.Show("Are you sure you want to cancel this transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    e.Cancel = True
                Else
                    e.Cancel = False
                    clearfields()
                    cbxSender.Select()
                    btnAdd.Text = "Record"
                    dgvRecv.Rows.Clear()
                    dgvRecv.Columns.Clear()
                End If
            Else
                e.Cancel = False
            End If

            With Me
                .dgvRecv.AllowUserToAddRows = True
                .dgvRecv.Enabled = True
            End With


        End If


    End Sub

    Private Sub tbxDocNum_Leave(sender As Object, e As EventArgs) Handles tbxDocNum.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Then
            validateDocuNumbers()
        Else

        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub CbxOwnership_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOwnership.SelectedIndexChanged

    End Sub

    Private Sub TbxDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxDocType.TextChanged
        doctypeselect()
    End Sub


    Private Sub doctypeselect()
        If tbxDocType.Text = "Receiving Report" Then

            q.loadSenderVendor(ComboBox1)
            Label9.Text = "Vendor :"
            Label18.Visible = True
            TextBox1.Visible = True
            tbxDocNum.Enabled = False
            q.fetchSeriesRecvAndDisplay()
            tbxDocNum.Text = series

        ElseIf tbxDocType.Text = "Goods Receipt Form" Then

            q.fetchreasonrec(ComboBox1)
            Label9.Text = "Vendor :"
            Label18.Visible = False
            TextBox1.Visible = False
            tbxDocNum.Enabled = False
            q.fetchGRFseries()
            tbxDocNum.Text = series

        Else

            q.fetchreasonrec(ComboBox1)
            Label9.Text = "Remarks :"
            Label18.Visible = False
            TextBox1.Visible = False
            tbxDocNum.Enabled = True
            tbxDocNum.Text = String.Empty
        End If
    End Sub
    Private Sub RecvGoodsMain_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub tbxDocNum_TextChanged(sender As Object, e As EventArgs) Handles tbxDocNum.TextChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)
        q.fetchIdDocType(tbxDocType.Text)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        y.FetchArea()
        q.fetchreferencename(docRefTypeId)

        If cbxSender.Text = "SUPPLIER" Then

            Dim from As String = ComboBox1.Text
            Dim area1 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series
            Dim cour As String = TextBox1.Text
            Dim refo As String = docCode + "-" + tbxRefDocNum.Text
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")

            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRecv.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString

                    datatable1.Rows.Add(datarow2)
                End If
            Next

            Dim reportDataSource As New ReportDataSource("DataSet1", datatable1)
            Print1.ReportViewer1.LocalReport.DataSources.Clear()
            Print1.ReportViewer1.LocalReport.DataSources.Add(reportDataSource)
            Print1.ReportViewer1.LocalReport.ReportPath = q.path + "Receiving Transactions\Listing\Report4.rdlc"

            Dim par As New ReportParameter("branch", area1)
            Print1.ReportViewer1.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print1.ReportViewer1.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print1.ReportViewer1.LocalReport.SetParameters(par2)

            Dim par3 As New ReportParameter("recv", from)
            Print1.ReportViewer1.LocalReport.SetParameters(par3)

            Dim par4 As New ReportParameter("Cour", cour)
            Print1.ReportViewer1.LocalReport.SetParameters(par4)

            Dim par5 As New ReportParameter("refno", refo)
            Print1.ReportViewer1.LocalReport.SetParameters(par5)


            Print1.ReportViewer1.RefreshReport()
            Print1.ShowDialog()

        Else

            Dim from As String
            Dim area1 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            If cbxSender.Text = String.Empty Then
                from = " "
            Else
                from = cbxSender.Text
            End If

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRecv.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString

                    datatable1.Rows.Add(datarow2)
                End If
            Next

            Dim reportDataSource As New ReportDataSource("DataSet1", datatable1)
            Print1.ReportViewer1.LocalReport.DataSources.Clear()
            Print1.ReportViewer1.LocalReport.DataSources.Add(reportDataSource)
            Print1.ReportViewer1.LocalReport.ReportPath = q.path + "Receiving Transactions\Listing\Report1.rdlc"

            Dim par As New ReportParameter("branch", area1)
            Print1.ReportViewer1.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print1.ReportViewer1.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print1.ReportViewer1.LocalReport.SetParameters(par2)

            Dim par3 As New ReportParameter("recv", "Received From:" + from)
            Print1.ReportViewer1.LocalReport.SetParameters(par3)

            Print1.ReportViewer1.RefreshReport()

            Print1.ShowDialog()

        End If
        btnAdd.PerformClick()

    End Sub
End Class