Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml

Public Class adjustInventoryRegister
    Private isManager As Boolean = False
    Private isAll As Boolean = True
    Private q As New qry
    Private areaId As String = ""



    Private Sub adjustInventoryRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpFrom.Value = Date.Now.ToShortDateString
        dtpTo.Value = Date.Now.ToShortDateString
        'dtpFrom.Format = DateTimePickerFormat.Short
        'dtpTo.Format = DateTimePickerFormat.Short
        q.loadAreaforUsers(cbxAreas)
        validateUser()
        newLVFormat()
        loadLV()

    End Sub

    Sub validateUser()
        Dim roleOfTheUser As String = ""
        roleOfTheUser = newHome.roleId
        If roleOfTheUser = "20019" Then

            rdAll.Visible = False
            rdSelectCateg.Visible = False
            isManager = False
        Else

            rdAll.Visible = True
            rdSelectCateg.Visible = True
            rdAll.Checked = True
            isManager = True
        End If

        hideCBXAreas()
    End Sub

    Private Sub hideCBXAreas()
        If rdAll.Checked = True Then
            isAll = True
            cbxAreas.SelectedIndex = -1
            cbxAreas.Visible = False
        ElseIf rdSelectCateg.Checked = True Then
            isAll = False
            cbxAreas.SelectedIndex = 0
            cbxAreas.Visible = True
        End If
    End Sub

    Private Sub newLVFormat()
        With lvInventoryAdjust
            .Alignment = ListViewAlignment.Top

            .Columns.Add("dummy", 0, HorizontalAlignment.Center) '0
            .Columns.Add("DOCUMENT DATE", 130, HorizontalAlignment.Center) '0
            .Columns.Add("ENTRY NUMBER", 180, HorizontalAlignment.Center) '1
            .Columns.Add("DOCUMENT NUMBER ", 180) '2
            .Columns.Add("REMARKS", 150)
            .Columns.Add("AREA", 180) '4
            .Columns.Add("ENCODED DATE", 150, HorizontalAlignment.Center) '5
            .Columns.Add("APPROVED BY", 200) '6
            .Columns.Add("encBY", 0)
        End With
    End Sub

    Sub loadLV()
        If isManager = True Then
            q.loadLVforInventoryAdjustManager(lvInventoryAdjust, tbxEntryNum.Text, isAll, areaId, dtpFrom.Value, dtpTo.Value)
        Else
            q.loadLVforInventoryAdjustNONManager(lvInventoryAdjust, tbxEntryNum.Text, newHome.areaId, dtpFrom.Value, dtpTo.Value)
        End If
        If lvInventoryAdjust.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
    End Sub

    Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
        hideCBXAreas()
    End Sub

    Private Sub rdSelectCateg_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCateg.CheckedChanged
        hideCBXAreas()
    End Sub

    Private Sub tbxEntryNum_TextChanged(sender As Object, e As EventArgs) Handles tbxEntryNum.TextChanged
        loadLV()
    End Sub

    Private Sub dtpFrom_ValueChanged(sender As Object, e As EventArgs) Handles dtpFrom.ValueChanged
        'Dim currentDate As DateTime = dtpFrom.Value
        'Dim previousDate As DateTime = currentDate.AddDays(-1)
        If dtpFrom.Value >= dtpTo.Value Then
            dtpTo.Value = dtpFrom.Value
        End If
        loadLV()

    End Sub

    Private Sub dtpTo_ValueChanged(sender As Object, e As EventArgs) Handles dtpTo.ValueChanged
        If dtpFrom.Value >= dtpTo.Value Then
            dtpTo.Value = dtpFrom.Value
        End If
        loadLV()
    End Sub

    Private Sub cbxAreas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxAreas.SelectedIndexChanged
        areaId = q.fetchSenderId(cbxAreas.Text)
        loadLV()
    End Sub

    Private Sub lvInventoryAdjust_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvInventoryAdjust.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = lvInventoryAdjust.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
        Home_Page.Focus()
        Home_Page.Select()
    End Sub

    Sub fetchDatas()
        If lvInventoryAdjust.Items.Count > 0 Then
            Dim x As Integer = -1
            Dim entryId As String
            Dim encoder As String
            x = CInt(lvInventoryAdjust.SelectedItems(0).Index)
            entryId = ""
            entryId = lvInventoryAdjust.Items(x).SubItems(2).Text
            encoder = lvInventoryAdjust.Items(x).SubItems(8).Text
            'MsgBox(entryId)
            With adjustInventoryRegisterDetails
                .id = entryId
                .encoder = encoder
                .loadlv()
                Home_Page.Enabled = False
                Me.Enabled = False
                .Show()
            End With
        End If
    End Sub

    Private Sub lvInventoryAdjust_DoubleClick(sender As Object, e As EventArgs) Handles lvInventoryAdjust.DoubleClick
        fetchDatas()
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
        Dim columncount As Integer = lvInventoryAdjust.Columns.Count - 1
        For i As Integer = 0 To columncount
            sheet.Cells(rowhead, colhead) = lvInventoryAdjust.Columns(i).Text
            colhead = colhead + 1
        Next

        For Each item As ListViewItem In lvInventoryAdjust.Items
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub adjustInventoryRegister_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        lvInventoryAdjust.Clear()
    End Sub

    Private Sub AdjustInventoryRegister_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class