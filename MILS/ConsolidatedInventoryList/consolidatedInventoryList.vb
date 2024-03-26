Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml

Public Class consolidatedInventoryList
    Private q As New qry

    Private Sub consolidatedInventoryList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            newLVFormat()
            loadLV()
            'loadTbxProdNum()
            'loadTbxProdDes()
            q.loadCBXforConsolidatedInv(cbxArea)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub newLVFormat()
        With lvConsoInv
            .Columns.Add("dummy", 0)
            .Columns.Add("PRODUCT NUMBER", 170, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 530)
            .Columns.Add("AREA", 250)
            .Columns.Add("TOTAL QUANTITY", 240, HorizontalAlignment.Center)
        End With
    End Sub

    Sub loadLV()
        q.loadConsolidateInventory(lvConsoInv, tbxProdNum.Text, tbxProdDes, cbxArea.Text)
        If lvConsoInv.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub





    'Sub loadTbxProdNum()
    '    With tbxProdNum
    '        .Text = ""
    '        .Refresh()
    '        .AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        .AutoCompleteSource = AutoCompleteSource.CustomSource
    '        Dim col As New AutoCompleteStringCollection()
    '        q.suggestPDNOinInventoryAdjust(col)
    '        .AutoCompleteCustomSource = col
    '    End With
    'End Sub

    'Sub loadTbxProdDes()
    '    With tbxProdDes
    '        .Text = ""
    '        .Refresh()
    '        .AutoCompleteMode = AutoCompleteMode.SuggestAppend
    '        .AutoCompleteSource = AutoCompleteSource.CustomSource
    '        Dim col As New AutoCompleteStringCollection()
    '        q.suggestPDDesinInventoryAdjust(col)
    '        .AutoCompleteCustomSource = col
    '    End With
    'End Sub

    'Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
    '    showCBXArea()
    '    loadLV()
    'End Sub

    Private Sub rdSelectCateg_CheckedChanged(sender As Object, e As EventArgs)
        loadLV()
    End Sub

    'Sub showCBXArea()
    '    Try
    '        If rdAll.Checked = True Then
    '            cbxArea.SelectedIndex = -1
    '            cbxArea.Visible = False
    '        ElseIf rdSelectCateg.Checked = True Then
    '            cbxArea.SelectedIndex = 0
    '            cbxArea.Visible = True
    '        End If
    '    Catch ex As Exception
    '        MsgBox(ex.Message)
    '    End Try
    'End Sub

    Private Sub tbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadLV()
    End Sub

    Private Sub tbxProdNum_TextChanged(sender As Object, e As EventArgs) Handles tbxProdNum.TextChanged
        loadLV()
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        loadLV()
    End Sub

    Private Sub lvConsoInv_DoubleClick(sender As Object, e As EventArgs) Handles lvConsoInv.DoubleClick
        If lvConsoInv.Items.Count > 0 Then
            Dim x As Integer = CInt(lvConsoInv.SelectedItems(0).Index)
            Dim id As String = ""
            Dim area As String = ""
            id = lvConsoInv.Items(x).SubItems(1).Text
            area = lvConsoInv.Items(x).SubItems(3).Text

            Me.Enabled = False
            With consolidatedInvDetails
                .prodId = ""
                .prodArea = ""
                .prodId = id
                .prodArea = area
                .loadLV()
                .Show()
            End With
        Else
        End If
    End Sub

    Private Sub lvConsoInv_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvConsoInv.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvConsoInv.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs)
        Me.Close()

    End Sub

    'EXPORT TO EXCEL FILE THE DATA FROM LISTVIEW
    Private Sub btnExport_Click(sender As Object, e As EventArgs)
        SaveFileDialog1.Title = "Save Excel File"
        SaveFileDialog1.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xslx"
        SaveFileDialog1.ShowDialog()

        'exit if no file selected
        If SaveFileDialog1.FileName = "" Then
            Exit Sub
        End If

        'create objects to interface to Excel
        Dim xls As New Excel.Application
        Dim book As Excel.Workbook
        Dim sheet As Excel.Worksheet

        'create a workbook and get reference to first worksheet
        xls.Workbooks.Add()
        book = xls.ActiveWorkbook
        sheet = book.ActiveSheet

        'step through rows and columns and copy data to worksheet
        Dim row As Integer = 2
        Dim col As Integer = 1

        Dim rowhead As Integer = 1
        Dim colhead As Integer = 1
        Dim columns As New List(Of String)
        Dim columncount As Integer = lvConsoInv.Columns.Count - 1
        For i As Integer = 0 To columncount
            sheet.Cells(rowhead, colhead) = lvConsoInv.Columns(i).Text
            colhead = colhead + 1
        Next

        For Each item As ListViewItem In lvConsoInv.Items
            For i As Integer = 0 To item.SubItems.Count - 1
                sheet.Cells(row, col) = item.SubItems(i).Text
                col = col + 1
            Next
            row += 1
            col = 1
        Next
        'save the workbook and clean up
        book.SaveAs(SaveFileDialog1.FileName)
        xls.Workbooks.Close()
        xls.Quit()
        releaseObject(sheet)
        releaseObject(book)
        releaseObject(xls)
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





    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            cbxArea.Items.Clear()
            cbxArea.Refresh()
            RadioButton2.Checked = False
            loadLV()
            cbxArea.Enabled = False
            cbxArea.Visible = False

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            cbxArea.Visible = True
            cbxArea.Enabled = True
            q.loadCBXforConsolidatedInv(cbxArea)
        End If
    End Sub

    Private Sub ConsolidatedInventoryList_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed


    End Sub

    Private Sub consolidatedInventoryList_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        lvConsoInv.Clear()
    End Sub

    Private Sub TbxProdNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdNum.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub ConsolidatedInventoryList_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class