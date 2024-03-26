Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml

Public Class inventoryGoods
    Private q As New qry
    Private id As String = ""
    Private Sub inventoryGoods_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpFrom.Value = Date.Now.ToShortDateString
        dtpTo.Value = Date.Now.ToShortDateString

        If Home_Page.roleId = 20016 Or Home_Page.roleId = "20017" Or Home_Page.roleId = "20018" Then
            gbArea.Visible = True
        Else
            gbArea.Visible = False
        End If
        q.loadAreaforUsers(cbxArea)
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvInventory
            .Columns.Add("dummy", 0)
            .Columns.Add("DOCUMENT DATE", 140, HorizontalAlignment.Center)
            .Columns.Add("ENTRY NUMBER", 150, HorizontalAlignment.Center)
            .Columns.Add("DOCUMENT NUMBER", 200)
            .Columns.Add("REMARKS", 170)
            .Columns.Add("AREA", 150)
            .Columns.Add("ENCODED DATE", 150, HorizontalAlignment.Center)
            .Columns.Add("USER", 200)
        End With
    End Sub

    Sub loadLV()
        q.loadAutoBuildProductsRegister(lvInventory, tbxEntryNum, tbxDocNum, dtpFrom.Value, dtpTo.Value, newHome.areaId, newHome.roleId, cbxArea.Text)
        If lvInventory.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub tbxEntryNum_TextChanged(sender As Object, e As EventArgs) Handles tbxEntryNum.TextChanged
        loadLV()
    End Sub

    Private Sub dtpFrom_ValueChanged(sender As Object, e As EventArgs) Handles dtpFrom.ValueChanged
        loadLV()
    End Sub

    Private Sub dtpTo_ValueChanged(sender As Object, e As EventArgs) Handles dtpTo.ValueChanged
        loadLV()
    End Sub

    Private Sub lvInventory_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvInventory.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvInventory.Columns(e.ColumnIndex).Width
    End Sub
    Sub fetchDatas()
        If lvInventory.SelectedItems.Count <> 0 Then
            Dim x As Integer = 0
            Dim selectedId As String = ""
            x = CInt(lvInventory.SelectedItems(0).Index)
            selectedId = ""
            selectedId = lvInventory.Items(x).SubItems(2).Text
            'MsgBox(selectedId)
            With inventoryGoodsDetails
                .id = selectedId
                .loadLV()
                Home_Page.Enabled = False
                Me.Enabled = False
                inventoryGoodsDetails.Show()
            End With
        End If
    End Sub

    Private Sub lvInventory_DoubleClick(sender As Object, e As EventArgs) Handles lvInventory.DoubleClick
        fetchDatas()
    End Sub

    Private Sub tbxDocNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxDocNum.KeyPress
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub tbxDocNum_TextChanged(sender As Object, e As EventArgs) Handles tbxDocNum.TextChanged
        loadLV()
    End Sub


    Private Sub rdAllArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdAllArea.CheckedChanged
        If rdAllArea.Checked = True Then
            cbxArea.Visible = False
            cbxArea.SelectedIndex = -1
        ElseIf rdSelectCategArea.Checked = True Then
            cbxArea.Visible = True
            cbxArea.SelectedIndex = 0
        End If
        loadLV()
    End Sub

    Private Sub rdSelectCategArea_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCategArea.CheckedChanged
        If rdAllArea.Checked = True Then
            cbxArea.Visible = False
            cbxArea.SelectedIndex = -1
        ElseIf rdSelectCategArea.Checked = True Then
            cbxArea.Visible = True
            cbxArea.SelectedIndex = 0
        End If
        loadLV()
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        loadLV()
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
        Dim columncount As Integer = lvInventory.Columns.Count - 1
        For i As Integer = 0 To columncount
            sheet.Cells(rowhead, colhead) = lvInventory.Columns(i).Text
            colhead = colhead + 1
        Next

        For Each item As ListViewItem In lvInventory.Items
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

    Private Sub inventoryGoods_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvInventory.Clear()
    End Sub

    Private Sub InventoryGoods_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class