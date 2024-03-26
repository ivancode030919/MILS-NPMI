Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class StatusReport

    Private n As New qryv3
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String
    Private Sub StatusReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        Validatedatas()
        cb1.Items.Clear()
        CrystalReportViewer1.ReportSource = Nothing
    End Sub

    Public Sub Validatedatas()
        CrystalReportViewer1.Refresh()
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument


        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        Dim t1 As String = ""

        If cb1.Text = "ACTIVE" Then
            t1 = "1"
        ElseIf cb1.Text = "IN-ACTIVE" Then
            t1 = "2"
        ElseIf cb1.Text = "DISCONTINUED" Then
            t1 = "3"
        Else
            t1 = ""
        End If


        rd.Load(My.Application.Info.DirectoryPath + "\iStatus.rpt")
        rd.SetDatabaseLogon(RPuser, RPpass)
        n.generateGoodsStatus(txt1.Text, txt2.Text, t1, rd)

        CrystalReportViewer1.ReportSource = rd
        CrystalReportViewer1.Refresh()
        CrystalReportViewer1.Zoom(80)
    End Sub

    Private Sub BtnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Validatedatas()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            cb1.Visible = False
            RadioButton2.Checked = False
            cb1.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            RadioButton1.Checked = False
            cb1.Visible = True
            cb1items()
        End If
    End Sub


    Private Sub cb1items()
        cb1.Items.AddRange({"ACTIVE", "IN-ACTIVE", "DISCONTINUED"})

    End Sub


    Private Sub clearfields()
        txt1.Text = ""
        txt2.Text = ""
        RadioButton1.Checked = True
    End Sub

    Private Sub StatusReport_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        clearfields()
    End Sub

    Private Sub BtnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        clearfields()
    End Sub

    Private Sub StatusReport_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class