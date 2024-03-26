Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml
Imports System.ComponentModel

Public Class releaseRegister
    Public docTypeId As String
    Private q As New qry
    Private isAll As Boolean = True
    Private receiverId As String
    Private transId As String
    Private dateType As Boolean = True

    Private Sub releaseRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpfrom.Value = Date.Now.ToShortDateString
        dtpTo.Value = Date.Now.ToShortDateString

        If newHome.roleId = "20016" Or newHome.roleId = "20020" Then
            'If Home_Page.roleId = 20001 Or Home_Page.roleId = "20001" Then
            gbArea.Visible = True
        Else
            gbArea.Visible = False
        End If
        newLvFormat()
        loadCBX()
        'isNotAll()
        loadLV()
    End Sub

    Sub newLvFormat()
        With lvRelsListing
            .Columns.Add("dummy", 0)
            .Columns.Add("DOCUMENT DATE", 120, HorizontalAlignment.Center)
            .Columns.Add("ENTRY NUMBER", 110, HorizontalAlignment.Center)
            .Columns.Add("DOCUMENT TYPE", 130, HorizontalAlignment.Left)
            .Columns.Add("DOCUMENT NUMBER", 140, HorizontalAlignment.Left)
            .Columns.Add("RELEASE TO", 200, HorizontalAlignment.Left)
            .Columns.Add("AREA", 170, HorizontalAlignment.Left)
            .Columns.Add("ENCODED DATE", 120, HorizontalAlignment.Center)
            .Columns.Add("USER", 200, HorizontalAlignment.Left)
            .Columns.Add("series", 0)


        End With
    End Sub

    Sub loadTBXdocType()
        With tbxDocType
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestDocTypeRelsListing(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Sub loadLV()
        q.loadLVRelsListing(lvRelsListing, cbxReceiver.Text, dtpfrom.Value, dtpTo.Value, isAll, docTypeId, tbxDocNum.Text, newHome.areaId, newHome.roleId, cbxArea.Text)
        If lvRelsListing.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No Data Found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Sub loadCBX()
        loadTBXdocType()
        q.loadSenderRegister(cbxReceiver)
        q.loadAreaforUsers(cbxArea)
    End Sub

    'Sub isNotAll()
    '    If rdAll.Checked = True Then
    '        cbxReceiver.Visible = False
    '        cbxReceiver.SelectedIndex = -1
    '        isAll = True
    '    ElseIf rdIndi.Checked = True Then
    '        cbxReceiver.Visible = True
    '        cbxReceiver.SelectedIndex = 0
    '        isAll = False
    '    End If
    'End Sub

    Private Sub lvRelsListing_DoubleClick(sender As Object, e As EventArgs) Handles lvRelsListing.DoubleClick
        If lvRelsListing.Items.Count > 0 Then
            Dim x As Integer = CInt(lvRelsListing.SelectedItems(0).Index)
            transId = ""
            transId = lvRelsListing.Items(x).SubItems(2).Text

            Me.Enabled = False
            With releaseDetails
                .transId = ""
                .transId = transId
                .series1 = lvRelsListing.Items(x).SubItems(9).Text
                .area1 = lvRelsListing.Items(x).SubItems(6).Text
                .loadDetails()
                .cbxReceiver.Select()
                .Show()
            End With
        Else
        End If
    End Sub

    Private Sub rdIndi_CheckedChanged(sender As Object, e As EventArgs) Handles rdIndi.CheckedChanged
        If rdIndi.Checked Then
            loadCBX()
            cbxReceiver.Visible = True
        End If

    End Sub

    Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
        If rdAll.Checked = True Then
            'cbxReceiver.Visible = False
            cbxReceiver.SelectedIndex = -1
            cbxReceiver.Text = ""
            cbxReceiver.Visible = False
        End If
        loadLV()
    End Sub

    Private Sub cbxReceiver_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxReceiver.SelectedIndexChanged
        receiverId = q.fetchSenderId(cbxReceiver.Text)
        If Not cbxReceiver.SelectedIndex = -1 Then
            loadLV()
        End If
    End Sub

    Private Sub tbxDocType_Leave(sender As Object, e As EventArgs) Handles tbxDocType.Leave
        q.fetchIdDocTypeListingRels(tbxDocType.Text)
        loadLV()
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
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

    Private Sub tbxDocNum_TextChanged(sender As Object, e As EventArgs) Handles tbxDocNum.TextChanged
        loadLV()
    End Sub

    Private Sub lvRelsListing_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvRelsListing.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvRelsListing.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub rdAllArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdAllArea.CheckedChanged
        If rdAllArea.Checked Then
            cbxArea.Visible = False
            cbxArea.SelectedIndex = -1
            cbxArea.Text = ""
        End If
    End Sub

    Private Sub rdSelectCategArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCategArea.CheckedChanged
        If rdSelectCategArea.Checked Then
            loadCBX()
            cbxArea.Visible = True
        End If
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

    Private Sub ReleaseRegister_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        lvRelsListing.Clear()
    End Sub

    Private Sub ReleaseRegister_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class