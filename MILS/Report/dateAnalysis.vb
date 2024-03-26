Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class dateAnalysis
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String
    Private Sub rdAll_CheckedChanged(sender As Object, e As EventArgs) Handles rdAll.CheckedChanged
        showCBX()

    End Sub

    Private Sub rdSelect_CheckedChanged(sender As Object, e As EventArgs) Handles rdSelect.CheckedChanged
        showCBX()
    End Sub

    Sub showCBX()
        If rdAll.Checked = True Then
            cbxCategory.SelectedIndex = -1
            cbxCategory.Visible = False
            cbxCategory.Items.Clear()
        ElseIf rdSelect.Checked = True Then
            loadCBXModuleType()
            cbxCategory.Visible = True
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

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Try
            validateFields()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub dateAnalysis_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        q.reportlog()
        loadCBXAreaType()
        crvDateAnalysis.ReportSource = Nothing
    End Sub

    Sub loadCBXModuleType()
        q.loadModuleTypesCBX(cbxCategory)
    End Sub

    Sub loadCBXAreaType()
        q.loadDateAnalysisArea(clbAreas)
    End Sub
    'MsgBox("'" + cbxCategory.Text + "%'")

    Private Sub validateFields()

        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument
        Dim selectedArea As String = clbAreas.Text
        Dim ModuleCategory As String = cbxCategory.Text
        Dim str As Date = dtpStart.Value.ToString
        Dim stp As Date = dtpStop.Value.ToString

        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        If RadioButton1.Checked Then
            rd.Load(My.Application.Info.DirectoryPath + "\dateallAnalysis.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateDateallAnalysis(str, stp, ModuleCategory, rd)

        ElseIf RadioButton2.Checked = True Then
            rd.Load(My.Application.Info.DirectoryPath + "\dateAnalysis.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateDateAnalysis(str, stp, ModuleCategory, selectedArea, rd)

        End If

        crvDateAnalysis.ReportSource = rd
        crvDateAnalysis.Zoom(150)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        clbAreas.Items.Clear()
        cbxCategory.Items.Clear()
        dtpStart.Value = Date.Now
        dtpStop.Value = Date.Now
        q.loadModuleTypesCBX(cbxCategory)

        rdAll.Checked = True
        rdSelect.Checked = False

        dtpStart.Select()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            clbAreas.Enabled = True
            clbAreas.Visible = True
            q.loadDateAnalysisArea(clbAreas)
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            clbAreas.Items.Clear()
            RadioButton2.Checked = False

            clbAreas.Enabled = False
            clbAreas.Visible = False
        End If
    End Sub

    Private Sub DateAnalysis_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
    End Sub

    Private Sub DateAnalysis_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class