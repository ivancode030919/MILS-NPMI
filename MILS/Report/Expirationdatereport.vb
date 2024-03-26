Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Expirationdatereport
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String

    Private Sub Expirationdatereport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadCBXAreaType()
        q.reportlog()
        Validatedatas()
        CrystalReportViewer1.ReportSource = Nothing

    End Sub

    Private Sub aging()
        If RadioButton1.Checked Then

        End If
    End Sub

    Private Sub Validatedatas()
        CrystalReportViewer1.Refresh()
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument

        Dim r1 As String
        Dim r2 As String
        Dim r3 As String = ComboBox1.Text
        Dim r4 As String = ""



        If radio1.Checked = True Then
            r4 = 1
        ElseIf radio2.Checked = True Then
            r4 = 0
        End If
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3
        If RadioButton1.Checked Then
            r1 = "0"
            r2 = "90"
            rd.Load(My.Application.Info.DirectoryPath + "\iExpirationReport.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateExpiration(txt1.Text, txt2.Text, r1, r2, r3, r4, rd)
        ElseIf RadioButton2.Checked Then
            r1 = "91"
            r2 = "180"
            rd.Load(My.Application.Info.DirectoryPath + "\iExpirationReport.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateExpiration(txt1.Text, txt2.Text, r1, r2, r3, r4, rd)
        ElseIf RadioButton3.Checked Then
            r1 = "181"
            r2 = "270"
            rd.Load(My.Application.Info.DirectoryPath + "\iExpirationReport.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateExpiration(txt1.Text, txt2.Text, r1, r2, r3, r4, rd)
        ElseIf RadioButton4.Checked Then
            r1 = "271"
            r2 = "10000000"
            rd.Load(My.Application.Info.DirectoryPath + "\iExpirationReport.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateExpiration(txt1.Text, txt2.Text, r1, r2, r3, r4, rd)

        ElseIf RadioButton7.Checked Then

            r1 = "-10000000"
            r2 = "0"
            r4 = ""
            rd.Load(My.Application.Info.DirectoryPath + "\iExpirationReport.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateExpiration(txt1.Text, txt2.Text, r1, r2, r3, r4, rd)


        End If

        CrystalReportViewer1.ReportSource = rd
        CrystalReportViewer1.Refresh()
        CrystalReportViewer1.Zoom(80)
    End Sub


    Private Sub BtnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Validatedatas()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Sub loadCBXAreaType()
        q.loadDateAnalysisArea(ComboBox1)
    End Sub

    Private Sub Radio1_CheckedChanged(sender As Object, e As EventArgs) Handles radio1.CheckedChanged
        If radio1.Checked = True Then
            radio2.Checked = False
        End If
    End Sub

    Private Sub Radio2_CheckedChanged(sender As Object, e As EventArgs) Handles radio2.CheckedChanged
        If radio2.Checked = True Then
            radio1.Checked = False
        End If
    End Sub

    Private Sub Expirationdatereport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
        clearfields()
        CrystalReportViewer1.Refresh()
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked Then
            ComboBox1.Visible = False
            RadioButton6.Checked = False
        End If

    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked Then
            ComboBox1.Visible = True
            RadioButton5.Checked = False
        End If
    End Sub

    Private Sub Txt1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '
        '
        '    ")
        'End If

    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearfields()
    End Sub

    Private Sub clearfields()
        radio2.Checked = True
        RadioButton1.Checked = True
        RadioButton5.Checked = True
        txt1.Text = ""
        txt2.Text = ""
    End Sub

    Private Sub Expirationdatereport_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class