Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class OPreport
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub OPreport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        RadioButton1.Checked = True
        Validatedatas()
        CrystalReportViewer1.ReportSource = Nothing
    End Sub


    Sub cbvalue()
        ComboBox1.Items.AddRange({"BUY", "HOLD"})
    End Sub

    Private Sub Validatedatas()
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        Dim r1 As String = "0"
        Dim r2 As String = "0"
        Dim r3 As String = ComboBox1.Text
        Dim r4 As String = TextBox1.Text
        Dim r5 As String = TextBox2.Text

        If RadioButton1.Checked Then
            r1 = "1"
            RadioButton2.Checked = False
            RadioButton5.Checked = False
        ElseIf RadioButton2.Checked Then
            r2 = "2"
            RadioButton1.Checked = False
            RadioButton5.Checked = False

        ElseIf RadioButton5.Checked Then
            RadioButton2.Checked = False
            RadioButton1.Checked = False
            r1 = "0"
            r2 = "0"
        Else
            RadioButton5.Checked = True
        End If
        rd.Load(My.Application.Info.DirectoryPath + "\iorderpoint.rpt")
        rd.SetDatabaseLogon(RPuser, RPpass)
        q.generateorderpoint(r4, r5, r1, r2, r3, rd)



        CrystalReportViewer1.ReportSource = rd
        CrystalReportViewer1.Refresh()
        CrystalReportViewer1.Zoom(80)
    End Sub

    Private Sub BtnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Validatedatas()
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked Then
            ComboBox1.Enabled = True
            cbvalue()
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            ComboBox1.Items.Clear()
            ComboBox1.Enabled = False
            ComboBox1.Text = ""
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearfields()
    End Sub



    Private Sub clearfields()
        TextBox1.Text = ""
        TextBox2.Text = ""
        RadioButton3.Checked = True
        RadioButton1.Checked = True
    End Sub

    Private Sub OPreport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        clearfields()
        CrystalReportViewer1.Refresh()
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub OPreport_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class