Imports Microsoft.Reporting.WinForms
Public Class releaseGoods
    Private q As New qry
    Private y As New qryv3
    'Private table As New DataTable("Table")
    Public docTypeId As String
    Public docRefTypeId As String
    Public receiverId As String
    Private slctedRow As Integer
    Public series1 As String = ""
    Public branch As String = ""
    Public area1 As String = ""

    Private Sub releaseGoods_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newDGVFormat()
        loadForm()
        dtpRefDate.MaxDate = DateTime.Now
    End Sub

    Sub loadForm()
        loadTbxDocType()
        loadtbxRefDocType()
        q.loadSender(cbxReceiver)
        cbxOwnership.SelectedIndex = 0
    End Sub

    Private Sub dgvRels_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRels.CellDoubleClick
        Dim row As Integer = dgvRels.CurrentCell.RowIndex
        If e.ColumnIndex = 1 Then


            With selectProducts
                .rel = True
                .loadLV()
                .Show()
                .targetRow1 = row
            End With
        End If
    End Sub

    Sub newDGVFormat()


        With dgvRels

            .AllowUserToAddRows = True

            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35

            .Columns.Add("0", "id")
            .Columns.Add("1", "PRODUCT NUMBER")
            .Columns.Add("2", "DESCRIPTION")
            .Columns.Add("3", "BATCH CODE")
            .Columns.Add("4", "LOCATION")
            .Columns.Add("5", "QUANTITY")
            .Columns.Add("6", "bId")
            .Columns.Add("7", "lId")
            .Columns.Add("8", "EXPIRATION DATE")

            .Columns(0).Visible = False
            .Columns(6).Visible = False
            .Columns(7).Visible = False

            .Columns(1).Width = 150
            .Columns(2).Width = 523
            .Columns(3).Width = 95
            .Columns(4).Width = 95
            .Columns(5).Width = 100
            .Columns(8).Width = 175

            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = False
            .Columns(6).ReadOnly = True
            .Columns(7).ReadOnly = True
            .Columns(8).ReadOnly = True


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 8, FontStyle.Bold)
            Next


        End With
    End Sub


    Private Sub dgvRels_Click(sender As Object, e As EventArgs) Handles dgvRels.Click
        If dgvRels.Rows.Count = 0 Then
            l5.Visible = True
            l5.ForeColor = Color.Red
            l5.Text = "Please select valid data..."
        Else
            If dgvRels.Rows.Count > 1 Then
                l5.Visible = True
                slctedRow = dgvRels.CurrentCell.RowIndex
            End If
        End If
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

    Private Sub cbxReceiver_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbxReceiver.SelectedValueChanged
        receiverId = ""
        receiverId = q.fetchSenderId(cbxReceiver.Text)
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub tbxRefDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxRefDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub tbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) Then
            'q.fetchIdDocType(tbxDocType.Text)
            q.fetchIdDocTypeRels(tbxDocType.Text)
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
            q.fetchIdRefDocTypeRels(tbxRefDocType.Text)
            'MsgBox(docRefTypeId)
        Else

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
                                    'dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                    e.Cancel = True
                                    l5.Visible = True
                                    l5.Text = "Duplicate data are not valid."
                                    repeatingId = True
                                Else
                                    'dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.White
                                End If
                            End If
                        Next
                    Next

                    Try
                        isValidQty = q.validationQuantityProd(idDGV, qtyDGV) ' check the quantity
                    Catch ex As Exception
                        MessageBox.Show("Fields with * are required.", "Error",
                                           MessageBoxButtons.OK, MessageBoxIcon.Error)
                        l1.Visible = True
                        l2.Visible = True
                        l3.Visible = True
                        l4.Visible = True
                        l5.Visible = True
                        Label5.Visible = True
                        Label7.Visible = True
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
                        l5.Visible = True
                    End If
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub releaseGoods_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If dgvRels.Rows.Count > 1 Then
            If e.KeyCode = Keys.Delete Then
                Dim result As Integer = MessageBox.Show("Are you sure want to remove selected row?", "Remove Details", MessageBoxButtons.YesNo)
                If result = DialogResult.No Then
                    'Exit Sub
                ElseIf result = DialogResult.Yes Then
                    Try
                        dgvRels.Rows.RemoveAt(slctedRow)
                    Catch ex As Exception
                        MsgBox(ex.Message)
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



    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If btnAdd.Text = "Record" Then

            If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else

                'If cbxReceiver.Text = "SUPPLIER" Then
                '    q.fetchrelPurchaseRetirnseries()
                'Else
                '    q.fetchSeriesRelAndDisplay()
                'End If

                validateFields()

            End If
        ElseIf btnAdd.Text = "Add New Entry" Then
            With Me
                .dgvRels.AllowUserToAddRows = True
                .dgvRels.Enabled = True
            End With
            clearfields()
            cbxReceiver.Select()
            Button1.Visible = True
            btnAdd.Text = "Record"
            l5.Visible = True
            newDGVFormat()
        End If
    End Sub

    Sub clearfields()
        cbxReceiver.SelectedIndex = -1
        tbxDocType.Text = ""
        docTypeId = ""
        tbxDocNum.Text = ""
        ComboBox1.Text = ""
        tbxEntry.Text = ""
        tbxRefDocType.Text = ""
        docRefTypeId = ""
        tbxRefDocNum.Text = ""
        cbxOwnership.SelectedIndex = 0

        dgvRels.Rows.Clear()
        dgvRels.Columns.Clear()

    End Sub

    Sub validateFields()
        Dim dgvCount As Integer = dgvRels.Rows.Count - 1
        If cbxReceiver.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or String.IsNullOrWhiteSpace(tbxDocNum.Text) Or String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or String.IsNullOrWhiteSpace(ComboBox1.Text) Or dgvCount = 0 Then
            MessageBox.Show("Fields with '*' are required", "Info:", MessageBoxButtons.OK)
            l1.Visible = True
            l2.Visible = True
            l3.Visible = True
            l4.Visible = True
            l5.Visible = True
            Label5.Visible = True
            Label7.Visible = True
            Exit Sub

        Else

            q.addRelsTransaction(docTypeId, tbxDocNum.Text, dtpRefDate.Value, newHome.userId, receiverId,
                                     ComboBox1.Text, cbxOwnership.Text, docRefTypeId, tbxRefDocNum.Text, newHome.areaId, series1)
        End If

    End Sub

    Private Sub releaseGoods_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If btnAdd.Text = "Add New Entry" Then
            e.Cancel = False

        Else
            If Not cbxReceiver.SelectedIndex = -1 Or Not String.IsNullOrWhiteSpace(tbxDocType.Text) _
            Or Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Or Not String.IsNullOrWhiteSpace(tbxRefDocType.Text) _
            Or dgvRels.Rows.Count > 0 Then
                If (MessageBox.Show("Are you sure you want to cancel this transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                    e.Cancel = True
                Else
                    e.Cancel = False
                    clearfields()
                End If
            Else
                e.Cancel = False
            End If

        End If

    End Sub

    Sub ValidateDocumentNumber()
        q.validateDocumentRelsTransaction(tbxDocType.Text, docTypeId, tbxDocNum.Text, newHome.areaId)
    End Sub

    Private Sub tbxDocNum_Leave(sender As Object, e As EventArgs) Handles tbxDocNum.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) And Not String.IsNullOrWhiteSpace(tbxDocNum.Text) Then
            ValidateDocumentNumber()
        End If
    End Sub

    Private Sub CbxReceiver_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxReceiver.SelectedIndexChanged
        DocSelection()
    End Sub

    Private Sub tbxDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxDocType.TextChanged
        DocSelection()
    End Sub


    Private Sub DocSelection()

        If cbxReceiver.Text = "OS CUSTOMER" Then
            tbxDocType.Text = "Delivery Receipt"
            ComboBox1.DropDownStyle = ComboBoxStyle.Simple
            Label9.Text = "Customer :"
            q.FetchSeriesinReleasing2()
            tbxDocNum.Text = series1
            tbxDocNum.Enabled = False

        ElseIf cbxReceiver.Text = "CUSTOMER" Then
            tbxDocType.Text = "Delivery Receipt"
            ComboBox1.DropDownStyle = ComboBoxStyle.Simple
            Label9.Text = "Remarks :"
            tbxDocNum.Text = ""
            tbxDocNum.Enabled = True

        ElseIf cbxReceiver.Text = "SUPPLIER" Then

            q.loadSenderVendor(ComboBox1)
            tbxDocType.Text = "Purchase Return"
            Label9.Text = " Vendor :"
            ComboBox1.DropDownStyle = ComboBoxStyle.DropDownList
            q.FetchSeriesinReleasing1()
            tbxDocNum.Text = series1
            tbxDocNum.Enabled = False

        ElseIf tbxDocType.Text = "Goods Issue Form" Then

            q.loadSender(ComboBox1)
            Label9.Text = "Remarks :"
            ComboBox1.DropDownStyle = ComboBoxStyle.Simple
            q.FetchSeriesinReleasing()
            tbxDocNum.Text = series1
            tbxDocNum.Enabled = False

        ElseIf cbxReceiver.Text = "COMPANY EXPENSE" Then
            tbxDocType.Text = "Company Expense"
            Label9.Text = "Remarks :"
            ComboBox1.DropDownStyle = ComboBoxStyle.Simple
            q.FetchSeriesinReleasingCE()
            tbxDocNum.Text = series1
            tbxDocNum.Enabled = False
        Else

            Label9.Text = "Remarks :"
            tbxDocNum.Enabled = True
            tbxDocNum.Text = ""
        End If
        q.fetchIdDocTypeRels(tbxDocType.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        y.FetchArea()
        If cbxReceiver.Text = "SUPPLIER" Then

            Dim remarks As String
            Dim area2 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")
            If ComboBox1.Text = String.Empty Then
                remarks = " "
            Else
                remarks = ComboBox1.Text
            End If
            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRels.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString

                    datatable1.Rows.Add(datarow2)
                End If
            Next


            Dim reportDataSource As New ReportDataSource("DataSet2", datatable1)
            Print2.ReportViewer2.LocalReport.DataSources.Clear()
            Print2.ReportViewer2.LocalReport.DataSources.Add(reportDataSource)
            'Purchase Return Form
            Print2.ReportViewer2.LocalReport.ReportPath = q.path + "Releasing Transactions\Release Register\Report3.rdlc"

            Dim par As New ReportParameter("branch", area2)
            Print2.ReportViewer2.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print2.ReportViewer2.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print2.ReportViewer2.LocalReport.SetParameters(par2)

            Dim par3 As New ReportParameter("remarks", "Supplier: " + remarks)
            Print2.ReportViewer2.LocalReport.SetParameters(par3)

            Print2.ReportViewer2.RefreshReport()
            Print2.pr = 2
            Print2.ShowDialog()

        ElseIf cbxReceiver.Text = "OS CUSTOMER" Then

            Dim rels As String = cbxReceiver.Text
            Dim area2 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRels.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString

                    datatable1.Rows.Add(datarow2)

                End If
            Next

            Dim reportDataSource As New ReportDataSource("DataSet2", datatable1)
            Print2.ReportViewer2.LocalReport.DataSources.Clear()
            Print2.ReportViewer2.LocalReport.DataSources.Add(reportDataSource)
            'Delivery Receipt
            Print2.ReportViewer2.LocalReport.ReportPath = q.path + "Releasing Transactions\Release Register\Report5.rdlc"

            Dim par As New ReportParameter("branch", area2)
            Print2.ReportViewer2.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print2.ReportViewer2.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print2.ReportViewer2.LocalReport.SetParameters(par2)

            'Dim par3 As New ReportParameter("remarks", "Issued To: " + rels)
            Dim par3 As New ReportParameter("remarks", rels)
            Print2.ReportViewer2.LocalReport.SetParameters(par3)

            Print2.ReportViewer2.RefreshReport()

            Print2.pr = 1
            Print2.ShowDialog()

        ElseIf cbxReceiver.Text = "COMPANY EXPENSE" Then

            Dim remark As String = ComboBox1.Text
            Dim area2 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRels.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString
                    datatable1.Rows.Add(datarow2)
                End If
            Next

            Dim reportDataSource As New ReportDataSource("DataSet2", datatable1)
            Print2.ReportViewer2.LocalReport.DataSources.Clear()
            Print2.ReportViewer2.LocalReport.DataSources.Add(reportDataSource)
            Print2.ReportViewer2.LocalReport.ReportPath = q.path + "Releasing Transactions\Release Register\Report6.rdlc"

            Dim par As New ReportParameter("branch", area2)
            Print2.ReportViewer2.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print2.ReportViewer2.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print2.ReportViewer2.LocalReport.SetParameters(par2)

            Dim par3 As New ReportParameter("remarks", remark)
            Print2.ReportViewer2.LocalReport.SetParameters(par3)

            Print2.ReportViewer2.RefreshReport()

            Print2.ShowDialog()

        Else

            Dim rels As String = cbxReceiver.Text
            Dim area2 As String = branch
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dgvRels.Rows
                If Not row.IsNewRow Then
                    Dim datarow2 As DataRow = datatable1.NewRow
                    datarow2("goodId") = row.Cells(1).Value.ToString
                    datarow2("goodDes") = row.Cells(2).Value.ToString
                    datarow2("qty") = row.Cells(5).Value.ToString

                    datatable1.Rows.Add(datarow2)

                End If
            Next

            Dim reportDataSource As New ReportDataSource("DataSet2", datatable1)
            Print2.ReportViewer2.LocalReport.DataSources.Clear()
            Print2.ReportViewer2.LocalReport.DataSources.Add(reportDataSource)
            Print2.ReportViewer2.LocalReport.ReportPath = q.path + "Releasing Transactions\Release Register\Report2.rdlc"

            Dim par As New ReportParameter("branch", area2)
            Print2.ReportViewer2.LocalReport.SetParameters(par)

            Dim par1 As New ReportParameter("date", currentDate)
            Print2.ReportViewer2.LocalReport.SetParameters(par1)

            Dim par2 As New ReportParameter("series", ser)
            Print2.ReportViewer2.LocalReport.SetParameters(par2)

            Dim par3 As New ReportParameter("Rel", "Issued To: " + rels)
            'Dim par3 As New ReportParameter("remarks", rels)
            Print2.ReportViewer2.LocalReport.SetParameters(par3)

            Print2.ReportViewer2.RefreshReport()

            Print2.pr = 1
            Print2.ShowDialog()

        End If

        btnAdd.PerformClick()
    End Sub
End Class