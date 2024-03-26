Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml
Public Class goodsRegister
    Private q As New qry
    Private dateType As Boolean = True
    Private selectedId As String = ""

    Private Sub goodsRegister_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        newLVFormat()
        loadLV()
    End Sub

    Sub newLVFormat()
        With lvGoodsRegister
            .Columns.Add("ENTRY NUMBER", 150)
            .Columns.Add("DATE", 150)
            .Columns.Add("USER", 150)
            .Columns.Add("REMARKS", 156)
            .Columns.Add("ENCODED DATE", 150)
        End With
    End Sub

    Sub loadLV()
        q.loadGoodsRegisterHeader(lvGoodsRegister, tbxEntryId, dtpFrom.Value, dtpTo.Value)
        If lvGoodsRegister.Items.Count = 0 Then
            lblError.Visible = True
            lblError.Text = "No data found..."
        Else
            lblError.Visible = False
        End If
    End Sub

    Private Sub tbxEntryId_TextChanged(sender As Object, e As EventArgs) Handles tbxEntryId.TextChanged
        loadLV()
    End Sub

    Private Sub dtpFrom_ValueChanged(sender As Object, e As EventArgs) Handles dtpFrom.ValueChanged
        If dtpFrom.Value > dtpTo.Value Then
            dtpTo.Value = dtpFrom.Value
        End If
        loadLV()
    End Sub

    Private Sub dtpTo_ValueChanged(sender As Object, e As EventArgs) Handles dtpTo.ValueChanged
        If dtpFrom.Value > dtpTo.Value Then
            dtpTo.Value = dtpFrom.Value
        End If
        loadLV()
    End Sub

    Private Sub lvGoodsRegister_DoubleClick(sender As Object, e As EventArgs) Handles lvGoodsRegister.DoubleClick
        If lvGoodsRegister.SelectedItems.Count <> 0 Then
            Dim x As Integer = 0
            x = CInt(lvGoodsRegister.SelectedItems(0).Index)
            selectedId = ""
            selectedId = lvGoodsRegister.Items(x).SubItems(0).Text
            With goodsRegisterDetails
                .id = selectedId
                .fetchData(selectedId)
                Home_Page.Enabled = False
                Me.Enabled = False
                goodsRegisterDetails.Show()
            End With
        End If
    End Sub

    Private Sub lvGoodsRegister_ColumnWidthChanging(sender As Object, e As ColumnWidthChangingEventArgs) Handles lvGoodsRegister.ColumnWidthChanging
        e.Cancel = True
        e.NewWidth = Me.lvGoodsRegister.Columns(e.ColumnIndex).Width
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub btnClose_Click_1(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
    'EXPORT TO EXCEL FILE THE DATA FROM LISTVIEW

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

    Private Sub goodsRegister_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        lvGoodsRegister.Clear()
    End Sub

    Private Sub GoodsRegister_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress

    End Sub

    Private Sub GoodsRegister_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class