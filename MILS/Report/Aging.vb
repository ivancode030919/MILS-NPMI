Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Public Class Aging
    Private n As New qryv3
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String

    Private Sub Aging_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        Validatedatas()
        CrystalReportViewer1.ReportSource = Nothing
    End Sub

    Private Sub Validatedatas()
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument

        Dim r1 As String
        Dim r2 As String
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3
        'If radio1.Checked = True Then
        '    r4 = 1
        'ElseIf radio2.Checked = True Then
        '    r4 = 0
        'End If

        If RadioButton1.Checked = True Then
            r1 = "0"
            r2 = "6"
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)
        ElseIf RadioButton2.Checked = True Then
            r1 = "7"
            r2 = "12"
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)
        ElseIf RadioButton3.Checked = True Then
            r1 = "13"
            r2 = "24"
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)
        ElseIf RadioButton4.Checked = True Then
            r1 = "25"
            r2 = "60"
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)

        ElseIf RadioButton5.Checked = True Then
            r1 = "61"
            r2 = "1000000"
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)

        Else
            r1 = ""
            r2 = ""
            rd.Load(My.Application.Info.DirectoryPath + "\iAging.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            n.generateExpiration(txt1.Text, txt2.Text, r1, r2, rd)
        End If

        CrystalReportViewer1.ReportSource = rd
        CrystalReportViewer1.Refresh()
        CrystalReportViewer1.Zoom(80)
    End Sub

    Private Sub BtnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Validatedatas()
    End Sub

    Private Sub Txt1_Click(sender As Object, e As EventArgs) Handles txt1.Click
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        RadioButton4.Checked = False
        RadioButton5.Checked = False
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        txt1.Text = ""
        txt2.Text = ""
    End Sub

    Private Sub Aging_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MyBase.KeyPress

    End Sub

    Private Sub Aging_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class