Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml

Public Class recvListing
    Public docTypeId As String
    Private q As New qry
    Private isAll As Boolean = True
    Private senderId As String
    Private transId As String
    Private dateType As Boolean = True

    Sub newLvFormat()

        With lvRecvListing
            .Columns.Add("dummy", 0, HorizontalAlignment.Center)
            .Columns.Add("DOCUMENT DATE", 120, HorizontalAlignment.Center)
            .Columns.Add("ENTRY NUMBER", 110, HorizontalAlignment.Center)
            .Columns.Add("DOCUMENT TYPE", 120)
            .Columns.Add("DOCUMENT NUMBER", 140)
            .Columns.Add("RECEIVED FROM", 200)
            .Columns.Add("AREA", 170)
            .Columns.Add("ENCODED DATE", 110, HorizontalAlignment.Center)
            .Columns.Add("USER", 200)
            .Columns.Add("series", 200)
            .Columns(9).Width = 0

        End With
    End Sub

    Sub loadLV()
        q.loadLVRecvListing(lvRecvListing, cbxSender.Text, isAll, dtpfrom.Value, dtpTo.Value, docTypeId, tbxDocNum.Text, newHome.areaId, newHome.roleId, cbxArea.Text)
        If lvRecvListing.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No Data Found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Sub loadCBX()
        loadTBXdocType()
        q.loadSenderRegister(cbxSender)
        q.loadAreaforUsers(cbxArea)
    End Sub

    Sub loadTBXdocType()
        With tbxDocType
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestDocTypeRecvListing(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Private Sub dtpfrom_ValueChanged(sender As Object, e As EventArgs) Handles dtpfrom.ValueChanged
        If dtpfrom.Value > dtpTo.Value Then
            dtpTo.Value = dtpfrom.Value
        End If
        loadLV()
    End Sub

    Private Sub dtpTo_ValueChanged(sender As Object, e As EventArgs) Handles dtpTo.ValueChanged
        If dtpfrom.Value > dtpTo.Value Then
            dtpTo.Value = dtpfrom.Value
        End If
        loadLV()
    End Sub

    Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
        If rdAll.Checked Then
            cbxSender.Text = ""
            cbxSender.Visible = False
            cbxSender.Items.Clear()
        End If
        loadLV()
    End Sub

    Private Sub rdIndi_CheckedChanged(sender As Object, e As EventArgs) Handles rdIndi.CheckedChanged
        If rdIndi.Checked Then
            senderId = q.fetchSenderId(cbxSender.Text)
            cbxSender.Visible = True

        End If
        loadLV()
        loadCBX()

    End Sub

    'Private Sub cbxSender_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSender.SelectedIndexChanged

    '    'MsgBox(senderId)
    '    If Not cbxSender.SelectedIndex = -1 Then
    '        loadLV()
    '    End If
    'End Sub

    Private Sub cbxOwnership_SelectedIndexChanged(sender As Object, e As EventArgs)
        loadLV()
    End Sub

    Private Sub lvRecvListing_DoubleClick(sender As Object, e As EventArgs) Handles lvRecvListing.DoubleClick

        If lvRecvListing.Items.Count > 0 Then
            Dim x As Integer = CInt(lvRecvListing.SelectedItems(0).Index)
            transId = ""
            transId = lvRecvListing.Items(x).SubItems(2).Text
            recvListingDetails.transId = ""
            recvListingDetails.transId = transId
            recvListingDetails.loadDetails()
            recvListingDetails.cbxSender.Select()

            Me.Enabled = False
            recvListingDetails.Show()

            With recvListingDetails
                .area = lvRecvListing.Items(x).SubItems(6).Text
                .series = lvRecvListing.Items(x).SubItems(9).Text
            End With
        Else

        End If

    End Sub

    Private Sub tbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        q.fetchIdDocTypeListing(tbxDocType.Text)
        loadLV()
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub tbxDocType_TextChanged(sender As Object, e As EventArgs) Handles tbxDocType.TextChanged
        loadLV()
    End Sub

    Private Sub tbxDocNum_TextChanged(sender As Object, e As EventArgs) Handles tbxDocNum.TextChanged
        loadLV()
    End Sub

    Private Sub lvRecvListing_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvRecvListing.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvRecvListing.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub rdAllArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdAllArea.CheckedChanged
        If rdAllArea.Checked = True Then
            cbxArea.Visible = False
            cbxArea.SelectedIndex = -1
        Else
            cbxArea.Visible = True
            cbxArea.SelectedIndex = 0
        End If
        loadLV()
    End Sub

    Private Sub rdSelectCategArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCategArea.CheckedChanged
        If rdAllArea.Checked = True Then
            cbxArea.Visible = False
            cbxArea.SelectedIndex = -1
        Else
            cbxArea.Visible = True
            cbxArea.SelectedIndex = 0
        End If
        loadLV()
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        loadLV()
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        'Release an automation object
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub  'FUNCTIONS END HERE


    Private Sub recvListing_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        dtpfrom.Value = Date.Now.ToShortDateString
        dtpTo.Value = Date.Now.ToShortDateString
        newLvFormat()
        loadCBX()

        If newHome.roleId = 2016 Or newHome.roleId = 20020 Then
            gbArea.Visible = True
        Else
            gbArea.Visible = False
        End If

    End Sub

    Private Sub RecvListing_FormClosing_1(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        lvRecvListing.Clear()
    End Sub

    Private Sub RecvListing_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub CbxSender_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxSender.SelectedIndexChanged
        loadLV()
    End Sub
End Class