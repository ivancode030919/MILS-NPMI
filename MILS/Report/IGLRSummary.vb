Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class IGLRSummary
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String
    Private Sub validateDate()
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs)
        tbxProdDes.Text = ""
        tbxProdId.Text = ""
        dtpStart.Value = Date.Now
        dtpStop.Value = Date.Now

        dtpStart.Select()
    End Sub


    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click

    End Sub

    Private Sub ListView2_Click(sender As Object, e As EventArgs) Handles ListView2.Click

    End Sub

    Sub validateFields()
        'Dim isChecked As Boolean = False
        'Dim selectedPDNO As String = ""
        ''Dim isCheckedArea As Boolean = False
        Dim selectedArea As String = ComboBox1.Text
        Dim selectedPDNO As String = ""
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        For Each item As ListViewItem In ListView2.Items
            Dim id As String = item.SubItems(0).Text

            selectedPDNO &= id & ","
        Next

        ' Remove the last comma
        If selectedPDNO.Length > 0 Then
            selectedPDNO = selectedPDNO.TrimEnd(","c)
        End If

        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument
        Dim str As Date = dtpStart.Value.ToString
        Dim stp As Date = dtpStop.Value.ToString

        'MsgBox(str + stp + selectedPDNO.ToString + rd.ToString)
        rd.Load(My.Application.Info.DirectoryPath + "\IGLRSummary.rpt")
        rd.SetDatabaseLogon(RPuser, RPpass)

        q.generateIGLRSum(selectedPDNO, str, stp, selectedArea, rd)

        crvIGLR.ReportSource = rd
        crvIGLR.Refresh()
        crvIGLR.Zoom(107)
    End Sub

    Sub loadInventoryIGReport()
        q.loadInventoryIGLR(ListView1, tbxProdDes.Text, tbxProdId.Text)

    End Sub

    Private Sub IGLRSummary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadInventoryIGReport()
        crvIGLR.ReportSource = Nothing
        q.reportlog()
        newLVformat()
        newLVformat1()
    End Sub

    Private Sub dtpStart_ValueChanged_1(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        validateDate()
    End Sub

    Private Sub dtpStop_ValueChanged_1(sender As Object, e As EventArgs) Handles dtpStop.ValueChanged
        validateDate()
    End Sub

    Private Sub tbxProdId_Leave_1(sender As Object, e As EventArgs) Handles tbxProdId.Leave
        loadInventoryIGReport()
    End Sub

    Private Sub tbxProdDes_Leave_1(sender As Object, e As EventArgs) Handles tbxProdDes.Leave
        loadInventoryIGReport()
    End Sub

    Private Sub tbxProdId_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles tbxProdId.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If

    End Sub

    Private Sub btnGenerate_Click_1(sender As Object, e As EventArgs) Handles btnGenerate.Click
        If ListView2.Items.Count = 0 Then
            MessageBox.Show("No Selected Products...")
        Else

            Try
                validateFields()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub btnClear_Click_1(sender As Object, e As EventArgs) Handles btnClear.Click
        tbxProdDes.Text = ""
        tbxProdId.Text = ""
        dtpStart.Value = Date.Now
        dtpStop.Value = Date.Now

        dtpStart.Select()
    End Sub

    Private Sub IGLRSummary_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
    End Sub


    Sub newLVformat()
        With ListView1
            .Columns.Add("Pd#", 50)
            .Columns.Add("Descrition", 500)

        End With
    End Sub

    Sub newLVformat1()
        With ListView2
            .Columns.Add("Pd#", 50)
            .Columns.Add("Descrition", 500)
        End With
    End Sub



    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        For Each selectedItem As ListViewItem In ListView1.SelectedItems
            ListView2.Items.Add(selectedItem.Clone())
        Next
    End Sub

    Private Sub ListView2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView2.SelectedIndexChanged
        If ListView2.SelectedItems.Count > 0 Then
            Dim item As ListViewItem = ListView2.SelectedItems(0)
            ListView2.Items.Remove(item)
        End If
    End Sub

    Private Sub TbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadInventoryIGReport()
    End Sub

    Private Sub TbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadInventoryIGReport()
    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        ComboBox1.Visible = False
        ComboBox1.Items.Clear()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        ComboBox1.Visible = True
        q.loadAreasInReport(ComboBox1)
    End Sub

    Private Sub IGLRSummary_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class