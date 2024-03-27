Imports Microsoft.Reporting.WinForms

Public Class releaseDetails
    Public transId As String
    Private q As New qry
    Private n As New qryv3
    Public doctype As String = ""
    Public refdoctype As String = ""
    Public receiver As String = ""
    Public loc As String = ""
    Public batch As String = ""
    Public area As String = ""

    Public area1 As String = ""
    Public series1 As String = ""
    Private Sub releaseDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'newLVFormat()
        loadCBX()
        loadDetails()
        cbxReceiver.Select()
        n.batchIdsss()
        n.locidsss()

        If newHome.roleId = 20020 Then
            GroupBox1.Enabled = True
            loadclass()
            loadTbxDocType()
            loadtbxRefDocType()
        Else
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            gb3.Enabled = False
            GroupBox4.Enabled = False
            'GroupBox4.Location = New Point(800, 8)
        End If
    End Sub

    Private Sub loadclass()
        n.doctype(tbxDocType.Text)
        n.doctype2(tbxRefDocType.Text)


    End Sub

    Sub loadTbxDocType()
        With tbxDocType
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
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            'q.suggestRefDocType(col)
            q.suggestRefDocTypeRels(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub
    Sub loadCBX()
        q.loadSender(cbxReceiver) 'load sender combox
    End Sub

    Sub loadDetails()
        q.loadLVRelsDetails(transId)
    End Sub

    Private Sub releaseDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        releaseRegister.Enabled = True

    End Sub

    Sub closethisform()
        releaseRegister.Enabled = True

        Me.Dispose()
    End Sub

    Private Sub releaseDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub



    Private Sub lvRecvDetails_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs)
        'e.Cancel = True
        'e.NewWidth = Me.lvRecvDetails.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub Rb2_CheckedChanged(sender As Object, e As EventArgs) Handles rb2.CheckedChanged
        If rb2.Checked = True Then
            GroupBox2.Enabled = False
            gb3.Enabled = False
            GroupBox4.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Rb1_CheckedChanged(sender As Object, e As EventArgs) Handles rb1.CheckedChanged
        If rb1.Checked = True Then
            GroupBox2.Enabled = True
            gb3.Enabled = True
            GroupBox4.Enabled = True
            Button1.Enabled = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'n.relupdate(tbxEntry.Text, doctype, tbxDocNum.Text, dtpRefDate.Value, tbxRemarks.Text, tbxRefDocType.Text, tbxRefDocNum.Text, cbxOwnership.Text)
        n.relupdate(tbxEntry.Text, tbxDocNum.Text, dtpRefDate.Value, tbxRemarks.Text, tbxRefDocNum.Text, cbxOwnership.Text, doctype, receiver, refdoctype)
    End Sub

    Private Sub Dg1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellDoubleClick
        Dim row As Integer = dg1.CurrentCell.RowIndex
        If e.ColumnIndex = 1 Then

            With selectProducts
                .ed = True
                .loadLV()
                .Show()
                .targetRow = row
            End With
        End If
    End Sub

    Private Sub Dg1_CellValidating(sender As Object, e As DataGridViewCellValidatingEventArgs) Handles dg1.CellValidating

    End Sub

    Private Sub CbxReceiver_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxReceiver.SelectedIndexChanged
        loaddocs()
        n.relreceiver(cbxReceiver.Text)
    End Sub

    Private Sub loaddocs()
        If cbxReceiver.Text = "OS Customer" Then
            Label9.Text = "Customer :"
        ElseIf tbxDocType.Text = "Goods Receipt Form" Then
            Label9.Text = "Satellite :"
        ElseIf tbxDocType.Text = "Purchase Return" Then
            Label9.Text = "Vendor :"
        End If
    End Sub
    Private Sub TbxDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxDocType.TextChanged
        loadclass()
    End Sub

    Private Sub TbxRefDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxRefDocType.TextChanged
        loadclass()
    End Sub

    Private Sub Dg1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dg1.RowValidating
        Dim row As DataGridViewRow = dg1.Rows(e.RowIndex)
        Dim isValidQty As Boolean = True
        Dim repeatingId As Boolean = False
        Dim isEmpty As Boolean

        Try
            With dg1
                If .IsCurrentRowDirty Then
                    Dim idDGV As String = ""
                    idDGV = row.Cells(7).Value.ToString
                    Dim qtyDGV As String = ""
                    qtyDGV = row.Cells(5).Value.ToString
                    Dim id As String = ""
                    id = row.Cells(0).Value.ToString


                    For Each cell As DataGridViewCell In row.Cells 'is null
                        If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                            e.Cancel = True
                            isEmpty = True
                        End If
                    Next

                    For i As Integer = 0 To Me.dg1.RowCount - 1
                        For j As Integer = 0 To Me.dg1.RowCount - 1
                            If i <> j Then
                                If dg1.Rows(i).Cells(0).Value = dg1.Rows(j).Cells(0).Value Then
                                    'dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                    e.Cancel = True
                                    'l5.Visible = True
                                    'l5.Text = "Duplicate data are not valid."
                                    MsgBox("Duplicate data are not valid.")
                                    repeatingId = True
                                Else
                                    'dgvRels.Rows(i).DefaultCellStyle.BackColor = Color.White
                                End If
                            End If
                        Next
                    Next


                    Try
                        isValidQty = q.validationQuantityProdedit(idDGV, qtyDGV, id) ' check the quantity
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
                        'l5.Visible = True
                    End If


                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If cbxReceiver.Text = "SUPPLIER" Then

            Dim remarks As String
            Dim area2 As String = area1.ToString
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")
            If tbxRemarks.Text = String.Empty Then
                remarks = " "
            Else
                remarks = tbxRemarks.Text
            End If
            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dg1.Rows
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
            Dim area2 As String = tbxRemarks.Text
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")


            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dg1.Rows
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
            'Dim par3 As New ReportParameter("remarks", rels)
            'Print2.ReportViewer2.LocalReport.SetParameters(par3)

            Print2.ReportViewer2.RefreshReport()

            Print2.pr = 1
            Print2.ShowDialog()

        Else

            Dim rels As String = cbxReceiver.Text
            Dim area2 As String = area1.ToString
            Dim currentDate As Date = Date.Today
            Dim ser As String = series1
            Dim datatable1 As DataTable
            Dim dataset As New DataSet("Dataset")

            datatable1 = New DataTable("Mydatatable")
            datatable1.Columns.Add("goodId")
            datatable1.Columns.Add("goodDes")
            datatable1.Columns.Add("qty")


            dataset.Tables.Add(datatable1)
            For Each row As DataGridViewRow In dg1.Rows
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

    End Sub
End Class