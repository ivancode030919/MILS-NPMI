Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml
Public Class inventoryView
    Private q As New qry

    Private Sub inventoryView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadLV()

        If newHome.roleId = 20020 Or newHome.roleId = 20021 Then
            newLVFormatamdin()
            GroupBox3.Visible = True
            Button1.Visible = True
        Else
            newLVFormat()
            Button1.Visible = False
            GroupBox3.Visible = False
        End If

    End Sub

    Sub newLVFormat()
        With lvInventory
            .Columns.Add("ID", 0)
            .Columns.Add("PRODUCT NUMBER", 230, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 750)
            .Columns.Add("TOTAL QUANTITY", 230, HorizontalAlignment.Center)
        End With
    End Sub


    Sub newLVFormatamdin()
        With lvInventory
            .Columns.Add("ID", 0)
            .Columns.Add("PRODUCT NUMBER", 220, HorizontalAlignment.Center)
            .Columns.Add("DESCRIPTION", 620)
            .Columns.Add("TOTAL QUANTITY", 230, HorizontalAlignment.Center)
            .Columns.Add("STATUS", 100, HorizontalAlignment.Center)

        End With
    End Sub

    Sub loadLV()
        Dim fltr As String = ""

        If ComboBox1.Text = "Inactive" Then
            fltr = "2"

        ElseIf ComboBox1.Text = "Active" Then
            fltr = "1"

        ElseIf ComboBox1.Text = "Discontinued" Then
            fltr = "3"
        End If
        q.loadLVInventory(lvInventory, newHome.areaId, tbxPDNO.Text, tbxProdDes.Text, fltr)
        If lvInventory.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Sub clearFields()
        tbxPDNO.Text = ""
        tbxProdDes.Text = ""
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs)
        clearFields()
        loadLV()
    End Sub

    Private Sub tbxPDNO_TextChanged(sender As Object, e As EventArgs) Handles tbxPDNO.TextChanged
        loadLV()
    End Sub

    Private Sub tbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadLV()
    End Sub

    Private Sub tbxBatch_TextChanged(sender As Object, e As EventArgs)
        loadLV()
    End Sub

    Private Sub tbxLoc_TextChanged(sender As Object, e As EventArgs)
        loadLV()
    End Sub

    Private Sub lvInventory_DoubleClick(sender As Object, e As EventArgs) Handles lvInventory.DoubleClick




        If lvInventory.Items.Count <> 0 Then
            Dim x As Integer = CInt(lvInventory.SelectedItems(0).Index)
            Dim id As String = ""
            id = lvInventory.Items(x).SubItems(1).Text
            'MsgBox(id)

            Me.Enabled = False
            With invViewDetails
                .id = ""
                .id = id
                .loadLV()
                .tbxBatch.Select()
                .Show()
            End With
         
        Else

        End If
    End Sub

    Private Sub lvInventory_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvInventory.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvInventory.Columns(e.ColumnIndex).Width
    End Sub



    'Private Sub BtnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click

    'End Sub


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
    '    Dim columncount As Integer = lvInventory.Columns.Count - 1
    '    For i As Integer = 0 To columncount
    '        sheet.Cells(rowhead, colhead) = lvInventory.Columns(i).Text
    '        colhead = colhead + 1
    '    Next

    '    For Each item As ListViewItem In lvInventory.Items
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

    '    BackgroundWorker1.RunWorkerAsync()
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

    Private Sub inventoryView_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        lvInventory.Clear()
    End Sub

    Private Sub TbxPDNO_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxPDNO.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '    ")
        'End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadLV()
    End Sub

    Private Sub LvInventory_Click(sender As Object, e As EventArgs) Handles lvInventory.Click
        If newHome.roleId = 20020 Or newHome.roleId = 20021 Then
            Button1.Enabled = True


        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If lvInventory.SelectedItems.Count = 0 Then
            MessageBox.Show("Please Select Item to Update Status")
        Else
            Dim selectedItem As ListViewItem = lvInventory.SelectedItems(0)
            Dim id As String = selectedItem.SubItems(1).Text
            Dim description As String = selectedItem.SubItems(2).Text

            Me.Enabled = False

            With itemstatus
                .TextBox1.Text = id
                .TextBox2.Text = description
                .ShowDialog()
            End With




            Me.Enabled = True ' Re-enable the main form after the Status form is closed
        End If



    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            ComboBox1.Visible = False
            ComboBox1.Text = ""
            ComboBox1.Items.Clear()
            loadLV()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            ComboBox1.Visible = True
            choice()
        End If
    End Sub

    Private Sub choice()
        ComboBox1.Items.AddRange({"Active", "Inactive", "Discontinued"})
    End Sub

    Private Sub InventoryView_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub



    'Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
    '    For i As Integer = 0 To 100
    '        BackgroundWorker1.ReportProgress(i)
    '        Threading.Thread.Sleep(i)
    '    Next
    'End Sub

    'Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
    '    ProgressBar1.Value = e.ProgressPercentage
    '    tbxLoading.Text = e.ProgressPercentage.ToString()
    'End Sub

    'Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
    '    tbxLoading.Text = " "
    '    ProgressBar1.Value = 0
    'End Sub


End Class