Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml
Imports System.ComponentModel
Imports System.Threading

'MAIN FORM VIEW
Public Class consolidatedInvLocationDetails
    Private q As New qry

    Private Sub consolidatedInvLocationDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            ListViewFormat()
            q.loadCBXforConsolidatedInv(cbxArea)
        Catch ex As Exception

        End Try
    End Sub
    'LIST VIEW FORMAT FOR COLUMNS
    Private Sub ListViewFormat()
        With lvConsoInv
            '.View = View.Details
            .Columns.Add("dummy", 0, HorizontalAlignment.Center)
            .Columns.Add("PRODUCT NUMBER", 120, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 490, HorizontalAlignment.Left)
            .Columns.Add("AREA", 150, HorizontalAlignment.Center)
            .Columns.Add("BATCH", 150, HorizontalAlignment.Center)
            .Columns.Add("LOCATION", 120, HorizontalAlignment.Center)
            .Columns.Add("TOTAL QUANTITY", 150, HorizontalAlignment.Center)

        End With
    End Sub
    'LOAD LIST VIEW OF ITEMS
    Sub loadListView()
        'If rdAll.Checked Then
        q.loadConsolidateInventory4(lvConsoInv, tbxProdNum.Text, tbxProdDes, tbxLocStartFrom, tbxBatch, cbxArea.Text)
        If lvConsoInv.Items.Count = 0 Then
            lblErr.Visible = True
            lblErr.Text = "No data found..."
        Else
            lblErr.Visible = False
        End If
        'Else
        '    q.loadConsolidateInventory2(lvConsoInv, tbxProdNum, tbxProdDes, tbxLocStartFrom, tbxBatch, cbxArea.Text)
        'End If

    End Sub
    'LOAD AREA WHEN SELECTED
    Sub comboboxArea()
        Try
            If rdAll.Checked = True Then
                cbxArea.SelectedIndex = -1
                cbxArea.Visible = False
            ElseIf rdSelectCateg.Checked = True Then
                cbxArea.SelectedIndex = 0
                cbxArea.Visible = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
        comboboxArea()
        loadListView()
    End Sub

    Private Sub rdSelectCateg_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelectCateg.CheckedChanged
        comboboxArea()
        loadListView()
    End Sub
    Private Sub tbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadListView()
    End Sub

    Private Sub tbxProdNum_TextChanged(sender As Object, e As EventArgs) Handles tbxProdNum.TextChanged
        loadListView()
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        loadListView()
    End Sub

    Private Sub lvConsoInv_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvConsoInv.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvConsoInv.Columns(e.ColumnIndex).Width
    End Sub



    'EXPORT TO EXCEL FILE THE DATA FROM LISTVIEW

    'Private Sub btnExport_Click(sender As Object, e As EventArgs)

    '    SaveFileDialog1.Title = "Save Excel File"
    '    SaveFileDialog1.Filter = "Excel files (*.xls)|*.xls|Excel Files (*.xlsx)|*.xslx"
    '    SaveFileDialog1.ShowDialog()

    '    'exit if no file selected
    '    If SaveFileDialog1.FileName = "" Then
    '        Exit Sub
    '    End If

    '    'create objects to interface to Excel
    '    Dim xls As New Excel.Application
    '    Dim book As Excel.Workbook
    '    Dim sheet As Excel.Worksheet

    '    'create a workbook and get reference to first worksheet
    '    xls.Workbooks.Add()
    '    book = xls.ActiveWorkbook
    '    sheet = book.ActiveSheet

    '    'step through rows and columns and copy data to worksheet
    '    Dim row As Integer = 2
    '    Dim col As Integer = 1

    '    Dim rowhead As Integer = 1
    '    Dim colhead As Integer = 1
    '    Dim columns As New List(Of String)
    '    Dim columncount As Integer = lvConsoInv.Columns.Count - 1
    '    For i As Integer = 0 To columncount
    '        sheet.Cells(rowhead, colhead) = lvConsoInv.Columns(i).Text
    '        colhead = colhead + 1
    '    Next

    '    For Each item As ListViewItem In lvConsoInv.Items
    '        For i As Integer = 0 To item.SubItems.Count - 1
    '            sheet.Cells(row, col) = item.SubItems(i).Text
    '            col = col + 1
    '        Next
    '        row += 1
    '        col = 1
    '    Next
    '    'save the workbook and clean up
    '    book.SaveAs(SaveFileDialog1.FileName)
    '    xls.Workbooks.Close()
    '    xls.Quit()
    '    releaseObject(sheet)
    '    releaseObject(book)
    '    releaseObject(xls)



    'End Sub
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

    Private Sub tbxLocStartFrom_TextChanged(sender As Object, e As EventArgs) Handles tbxLocStartFrom.TextChanged
        loadListView()
    End Sub

    Private Sub tbxBatch_TextChanged(sender As Object, e As EventArgs) Handles tbxBatch.TextChanged
        loadListView()
    End Sub

    Private Sub consolidatedInvLocationDetails_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

    End Sub

    Private Sub TbxProdNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdNum.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        'End If
    End Sub

    Private Sub ConsolidatedInvLocationDetails_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub



    'SHOW PROGRESS BAR AT THE BUTTOM WHEN EXPORTING FILES TO EXCEL 
    'Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    '    Dim sum As Integer = 1
    '    For i As Integer = 1 To 100
    '        Thread.Sleep(200)
    '        'sum = sum
    '        BackgroundWorker1.ReportProgress(i)

    '        If BackgroundWorker1.CancellationPending Then
    '            e.Cancel = True
    '            BackgroundWorker1.ReportProgress(1)
    '            Return
    '        End If
    '    Next
    '    e.Result = sum
    'End Sub

    'Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
    '    ProgressBar1.Value = e.ProgressPercentage
    '    Label2.Text = e.ProgressPercentage.ToString() + "%"
    'End Sub

    'Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '    MsgBox("Data Successfully Exported!")
    '    If e.Cancelled Then
    '        Label2.Text = "Cancelled!"
    '    ElseIf e.Error IsNot Nothing Then
    '        Label2.Text = e.Error.Message
    '    Else
    '        Label2.Text = e.Result.ToString
    '    End If
    '    ProgressBar1.Value = 0
    '    Label2.Text = 0
    'End Sub






End Class