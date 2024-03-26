Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class IGLRSpecific
    Private q As New qry
    Dim selectingBatch As Boolean = False
    Dim selectingLoc As Boolean = False
    Public RPuser3 As String
    Public RPpass3 As String

    Private Sub IGLRSpecific_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        loadtbxpdDes()
        loadtbxProdNum()
        crvIGLR.ReportSource = Nothing

    End Sub

    Private Sub dtpStart_ValueChanged(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
        loadCBX()
    End Sub

    Private Sub dtpStop_ValueChanged(sender As Object, e As EventArgs) Handles dtpStop.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
        loadCBX()
    End Sub

    Sub loadtbxProdNum()
        With tbxProdNum
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestPDNOinInventoryAdjust(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Sub loadtbxpdDes()
        With tbxProdDes
            .Text = ""
            .Refresh()
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.CustomSource
            Dim col As New AutoCompleteStringCollection()
            q.suggestPDDesinInventoryAdjust(col)
            .AutoCompleteCustomSource = col
        End With
    End Sub

    Private Sub tbxProdNum_Leave(sender As Object, e As EventArgs) Handles tbxProdNum.Leave
        If Not String.IsNullOrWhiteSpace(tbxProdDes.Text) Then
            tbxProdDes.Text = ""

        End If

        q.fetchProductIDandDesIGLRSpecific(tbxProdNum.Text, tbxProdDes.Text)

        loadCBX()
        GroupBox4.Select()
    End Sub

    Private Sub tbxProdDes_Leave(sender As Object, e As EventArgs) Handles tbxProdDes.Leave
        If Not String.IsNullOrWhiteSpace(tbxProdNum.Text) Then
            tbxProdNum.Text = ""

        End If

        q.fetchProductIDandDesIGLRSpecific(tbxProdNum.Text, tbxProdDes.Text)

        loadCBX()
        GroupBox4.Select()
    End Sub

    Sub loadCBX()
        If Not String.IsNullOrWhiteSpace(tbxProdNum.Text) Or Not String.IsNullOrWhiteSpace(tbxProdDes.Text) Then
            q.FetchDataINCBXIGLRSpecific(tbxProdNum.Text, cbxArea, dtpStart.Value, dtpStop.Value)
        Else
            'MessageBox.Show("Please select Product Number/Description first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            cbxArea.SelectedIndex = -1
            cbxBatch.SelectedIndex = -1
            cbxLoc.SelectedIndex = -1
            cbxArea.Enabled = False
            cbxBatch.Enabled = False
            cbxLoc.Enabled = False
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        dtpStart.Value = Date.Now
        dtpStop.Value = Date.Now
        tbxProdNum.Text = ""
        tbxProdDes.Text = ""
        cbxArea.SelectedIndex = -1
        cbxBatch.SelectedIndex = -1
        cbxLoc.SelectedIndex = -1
        dtpStart.Select()
    End Sub

    Private Sub tbxProdNum_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxProdNum.KeyPress

    End Sub

    Sub loadReport()
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument
        Dim str As Date = dtpStart.Value.ToString
        Dim stp As Date = dtpStop.Value.ToString
        Dim pdNum As String = tbxProdNum.Text
        Dim area As String = cbxArea.Text
        Dim batch As String = cbxBatch.Text
        Dim loc As String = cbxLoc.Text

        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        rd.Load(My.Application.Info.DirectoryPath + "\IGLRSpecific.rpt")
        'rd.SetDatabaseLogon("sa", "p@ssw0rd")
        rd.SetDatabaseLogon(RPuser, RPpass)
        q.generateIGLRSpecific(pdNum, str, stp, area, batch, loc, rd)

        crvIGLR.ReportSource = rd
        crvIGLR.Refresh()
        crvIGLR.Zoom(83)
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If String.IsNullOrWhiteSpace(tbxProdNum.Text) Then
            MessageBox.Show("Please select Product Number/Description first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            Try
                loadReport()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub cbxArea_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxArea.SelectedIndexChanged
        q.fetchBatchAndLocationINGLSpecificCBX(tbxProdNum.Text, cbxArea.Text, cbxBatch, cbxLoc, dtpStart.Value, dtpStop.Value)
    End Sub

    Private Sub cbxBatch_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxBatch.SelectedIndexChanged
        'If selectingBatch = False Then
        q.fetchCBXLocationInLedgerReportSpecific(tbxProdNum.Text, cbxArea.Text, cbxBatch.Text, dtpStart.Value, dtpStop.Value, cbxLoc)
        'Else
        '    selectingBatch = Not selectingBatch
        '    Exit Sub
        'End If
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs)
        loadtbxpdDes()
    End Sub

    Private Sub IGLRSpecific_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
    End Sub

    Private Sub TbxProdNum_TextChanged(sender As Object, e As EventArgs) Handles tbxProdNum.TextChanged

    End Sub

    Private Sub IGLRSpecific_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    'Private Sub tbxProdNum_TextChanged(sender As Object, e As EventArgs) Handles tbxProdNum.TextChanged
    '    tbxProdDes.Text = ""
    'End Sub

    'Private Sub tbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged

    'End Sub
End Class