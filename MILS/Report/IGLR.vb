Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class IGLR
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String
    Private Sub IGLR_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        crvIGLR.ReportSource = Nothing
    End Sub

    Sub validateFields()
        'Dim isChecked As Boolean = False
        'Dim selectedPDNO As String = ""
        ''Dim isCheckedArea As Boolean = False
        Dim selectedArea As String = ComboBox1.Text
        Dim selectedPDNO As String = ""

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
        'Dim prodDes As String = tbxProdDes.Text

        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        If RadioButton1.Checked = True Then

            rd.Load(My.Application.Info.DirectoryPath + "\IGLRall.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateIGLRall(selectedPDNO, str, stp, rd)

        Else

            rd.Load(My.Application.Info.DirectoryPath + "\IGLR.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateIGLR(selectedPDNO, str, stp, selectedArea, rd)

        End If



        crvIGLR.ReportSource = rd
        crvIGLR.Refresh()
        crvIGLR.Zoom(83)
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        tbxProdDes.Text = ""
        tbxProdId.Text = ""
        dtpStart.Value = Date.Now
        dtpStop.Value = Date.Now
        ComboBox1.Text = ""
        ListView2.Items.Clear()
        dtpStart.Select()

    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
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

    Private Sub dtpStart_ValueChanged(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
    End Sub

    Private Sub dtpStop_ValueChanged(sender As Object, e As EventArgs) Handles dtpStop.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
    End Sub

    Sub loadInventoryIGReportr()
        q.loadInventoryIGLRr(ListView1, tbxProdDes.Text, tbxProdId.Text)
        If ListView1.Items.Count = 0 Then
            btnGenerate.Enabled = False
        Else
            btnGenerate.Enabled = True
        End If
    End Sub

    Private Sub TbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadInventoryIGReportr()

    End Sub

    Private Sub TbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadInventoryIGReportr()
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



    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
        'If e.IsSelected Then
        '    Dim newItem As New ListViewItem(e.Item.Text)
        '    ListView2.Items.Add(newItem)
        'End If
    End Sub
    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        ' Clear existing items in ListView2
        'ListView2.Items.Clear()

        ' Copy selected item(s) from ListView1 to ListView2
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

    Private Sub ListView2_Click(sender As Object, e As EventArgs) Handles ListView2.Click

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If String.IsNullOrWhiteSpace(ComboBox1.SelectedItem?.ToString()) Then
            GroupBox4.Enabled = False
            GroupBox2.Enabled = False
            ListView2.Enabled = False
            ListView1.Enabled = False
        Else
            GroupBox4.Enabled = True
            GroupBox2.Enabled = True
            ListView2.Enabled = True
            ListView1.Enabled = True
            loadInventoryIGReportr()
            newLVformat()
            newLVformat1()
        End If
    End Sub
    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            ComboBox1.Enabled = False
            ComboBox1.Visible = False
            ComboBox1.Items.Clear()
            GroupBox4.Enabled = True
            GroupBox2.Enabled = True
            ListView2.Enabled = True
            ListView1.Enabled = True
            loadInventoryIGReportr()
            newLVformat()
            newLVformat1()

        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            ComboBox1.Enabled = True
            ComboBox1.Visible = True
            q.loadAreasInReportr(ComboBox1)
        End If
    End Sub

    Private Sub IGLR_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        tbxProdId.Text = String.Empty
    End Sub

    Private Sub ListView1_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListView1.ColumnClick

    End Sub

    Private Sub IGLR_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class