Imports Microsoft.Reporting.WinForms

Public Class recvListingDetails
    Public transId As String
    Private q As New qry
    Private n As New qryv3
    Public senderId As String
    Public docTypeId As String = ""
    Public docCode As String = ""
    Public docRefTypeId As String = ""
    Public area As String = ""
    Public series As String = ""


    Private Sub recvListingDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loads()
        loadDetails()
        q.fetchreferencename(docRefTypeId)
        If newHome.roleId = 20020 Then
            GroupBox3.Visible = True

        Else
            GroupBox3.Visible = False
            'GroupBox1.Location = New Point(778, 9)
        End If

    End Sub


    Sub loadCBX()
        q.loadSender(cbxSender) 'load sender combox
    End Sub

    Public Sub loadDetails()
        q.loadLVRecvDetail(dg1, transId)
        n.doctype(tbxDocType.Text)
        n.doctype2(tbxRefDocType.Text)
    End Sub



    Private Sub recvListingDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        closethisform()
    End Sub

    Sub closethisform()
        recvListing.Enabled = True

        Me.Dispose()
    End Sub

    Private Sub recvListingDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs)
        closethisform()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles rb1.CheckedChanged
        If rb1.Checked Then
            GroupBox4.Enabled = True
            GroupBox1.Enabled = True
            dg1.Enabled = True
            rb2.Checked = False
            Button1.Enabled = True
        End If
    End Sub



    Private Sub manager()
        If newHome.roleId = 20020 Then
        Else
            GroupBox4.Enabled = False
            GroupBox1.Enabled = False
            GroupBox3.Enabled = False
            dg1.Enabled = False
            Button1.Enabled = False
        End If
    End Sub

    Private Sub Rb2_CheckedChanged(sender As Object, e As EventArgs) Handles rb2.CheckedChanged
        If rb2.Checked Then
            GroupBox4.Enabled = False
            GroupBox1.Enabled = False
            dg1.Enabled = False
            rb1.Checked = False
            Button1.Enabled = False
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        n.recvupdate(tbxEntry.Text, senderId, docTypeId, tbxDocNum.Text, dtpRefDate.Value, tbxRemarks.Text, docRefTypeId, tbxRefDocNum.Text, cbxOwnership.Text)
    End Sub


    Private Sub loads()
        loadTbxDocType()
        loadCBX()
        loadDetails()
        cbxSender.Select()
        manager()
        loadtbxRefDocType()
        q.fetchSenderId(cbxSender.Text)
    End Sub

    Sub loadTbxDocType()
        With tbxDocType

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

            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestRefDocType(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Private Sub CbxSender_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSender.SelectedIndexChanged
        senderId = q.fetchSenderId(cbxSender.Text)
    End Sub

    Private Sub TbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxDocType.Text) Then
            q.fetchIdDocType(tbxDocType.Text)
            'MsgBox(docTypeId)
        Else

        End If
    End Sub

    Private Sub TbxRefDocType_Leave(sender As Object, e As EventArgs) Handles tbxRefDocType.Leave
        If Not String.IsNullOrWhiteSpace(tbxRefDocType.Text) Then
            q.fetchIdRefDocType(tbxRefDocType.Text)
        Else

        End If
    End Sub

    Private Sub Dg1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellDoubleClick
        Dim row As Integer = dg1.CurrentCell.RowIndex
        If e.ColumnIndex = 1 Then

            Me.Enabled = False
            With recvSelectGood
                .ed = True
                .loadLV()
                .rowToEdit = row
                .ShowDialog()
                .tbxFilter.Select()
                .Focus()
            End With
        ElseIf e.ColumnIndex = 3 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With recvSelectBatch
                .ed = True
                .rowToEdit = row
                .Show()
                .tbxFilter.Select()
                .Focus()
            End With
        ElseIf e.ColumnIndex = 4 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With recvSelectLocations
                .ed = True
                .rowToEdit = row
                .Show()
                .tbxFilter.Select()
                .Focus()
            End With

        ElseIf e.ColumnIndex = 6 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With selectExpiration
                .ed = True
                .rowToEdit = row
                .Show()
                .Focus()
            End With
        End If
        Me.Enabled = True
    End Sub

    Private Sub Dg1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellContentClick

    End Sub

    Private Sub Dg1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellClick

    End Sub

    Private Sub Dg1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dg1.RowValidating

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If cbxSender.Text = "SUPPLIER" Then
            Dim from As String = tbxRemarks.Text
            Dim area1 As String = area.ToString
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
            For Each row As DataGridViewRow In dg1.Rows
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
            Dim area1 As String = area.ToString
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
            For Each row As DataGridViewRow In dg1.Rows
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

    End Sub

    Private Sub tbxDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxDocType.TextChanged
        If tbxDocType.Text = "Receiving Report" Then
            Label9.Text = "Vendor :"
        Else
            Label9.Text = "Remarks :"
        End If
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs)

    End Sub
End Class