Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared


Public Class Export
    Public RPuser3 As String
    Public RPpass3 As String
    Private q As New qry
    Private Sub Export_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CrystalReportViewer1.ReportSource = Nothing

        q.reportlog()
        cbvalue()
        loadall()
        loadCBX()
        dtpStart.Value = Date.Now.ToShortDateString
        dtpStop.Value = Date.Now.ToShortDateString
        CrystalReportViewer1.Refresh()
        cb1.Text = ""
        validation()
    End Sub

    Sub loadCBX()
        q.loadSenderRegister(cbx1)
    End Sub
    Sub loadall()
        q.loadAreaforUsers(ComboBox2)
    End Sub
    Sub cbvalue()
        cb1.Items.AddRange({"Receive Register",
                           "Release Register",
                             "Transfer Register",
                               "Count Register",
                                 "Build Inventory Register",
                                 "Item Register",
                                "Product Register",
                                "Inventory List",
                                 "Inventory List by Location"})
    End Sub

    Sub validation()

        If cb1.Text = "Receive Register" Then
            GroupBox1.Enabled = True
            txt1.Visible = True
            txt2.Visible = True
            GroupBox5.Enabled = True
            l1.Visible = True
            l2.Visible = True
            l1.Text = "Document Number :"
            l2.Text = "Document Type :"
            loadCBX()
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            GroupBox5.Visible = True
            GroupBox3.Enabled = True
            GroupBox4.Visible = True
            GroupBox4.Size = New Size(225, 120)
            GroupBox5.Location = New Point(7, 363)
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Release Register" Then
            GroupBox1.Enabled = True
            txt1.Visible = True
            txt2.Visible = True
            GroupBox5.Enabled = True
            l1.Visible = True
            l2.Visible = True
            l1.Text = "Document Number :"
            l2.Text = "Document Type :"
            loadCBX()
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            GroupBox4.Visible = True
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            GroupBox5.Visible = True
            GroupBox3.Enabled = True
            txt4.Visible = False
            GroupBox4.Size = New Size(225, 120)
            GroupBox5.Location = New Point(7, 363)
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Build Inventory Register" Then
            GroupBox4.Size = New Size(225, 120)
            txt1.Visible = True
            txt2.Visible = True
            GroupBox1.Enabled = True
            l1.Visible = True
            l2.Visible = True
            l3.Visible = False
            l1.Text = "Document Number :"
            l2.Text = "Entry Number :"
            GroupBox5.Visible = False
            GroupBox4.Visible = True
            GroupBox3.Enabled = True
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Item Register" Then
            GroupBox4.Size = New Size(225, 75)
            l1.Text = "Entry Number :"
            l1.Visible = True
            l2.Visible = False
            txt2.Visible = False
            txt1.Visible = True
            GroupBox3.Enabled = False
            GroupBox4.Visible = True
            GroupBox1.Enabled = True
            GroupBox5.Visible = False
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Transfer Register" Then

            l1.Text = "Document Number :"
            l2.Text = "Entry Number :"
            txt1.Visible = True
            txt2.Visible = True
            l1.Visible = True
            l2.Visible = True
            GroupBox5.Visible = False
            GroupBox4.Visible = True
            GroupBox3.Enabled = True
            GroupBox1.Enabled = True
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            GroupBox4.Size = New Size(225, 120)
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Count Register" Then
            GroupBox4.Size = New Size(225, 75)
            l1.Text = "Entry Number :"
            l2.Visible = False
            l1.Visible = True
            txt2.Visible = False
            txt1.Visible = True
            GroupBox3.Enabled = True
            GroupBox4.Visible = True
            GroupBox1.Enabled = True
            GroupBox5.Visible = False
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Product Register" Then
            GroupBox4.Size = New Size(225, 120)
            GroupBox1.Enabled = False
            l1.Text = "Product Number :"
            l2.Text = "Description :"
            txt1.Visible = True
            txt2.Visible = True
            l1.Visible = True
            l2.Visible = True
            GroupBox5.Visible = False
            GroupBox4.Visible = True
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            GroupBox3.Enabled = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Inventory List by Location" Then
            GroupBox4.Size = New Size(225, 220)
            l1.Text = "Product Number :"
            l2.Text = "Description :"
            l3.Text = "Location :"
            l4.Text = "Batch :"
            txt1.Visible = True
            txt2.Visible = True
            txt3.Visible = True
            txt4.Visible = True
            l1.Visible = True
            l2.Visible = True
            l3.Visible = True
            l4.Visible = True
            GroupBox5.Visible = False
            GroupBox3.Enabled = True
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            GroupBox4.Visible = True
            GroupBox1.Enabled = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True
        ElseIf cb1.Text = "Inventory List" Then
            GroupBox4.Size = New Size(225, 120)
            l1.Text = "Product Number :"
            l2.Text = "Description :"
            txt1.Visible = True
            txt2.Visible = True
            l1.Visible = True
            l2.Visible = True
            GroupBox3.Enabled = True
            GroupBox5.Visible = False
            GroupBox1.Visible = True
            GroupBox3.Visible = True
            GroupBox1.Enabled = False
            GroupBox4.Visible = True
            l3.Visible = False
            l4.Visible = False
            txt3.Visible = False
            txt4.Visible = False
            txt1.Text = ""
            txt2.Text = ""
            txt3.Text = ""
            txt4.Text = ""
            RadioButton1.Checked = True
            Button1.Enabled = True

        ElseIf cb1.Text = String.Empty Then

            GroupBox3.Enabled = False
            GroupBox5.Visible = False
            GroupBox1.Visible = False
            GroupBox3.Visible = False
            GroupBox1.Enabled = False
            GroupBox4.Visible = False

        End If

    End Sub

    Private Sub Cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        validation()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            ComboBox2.Text = ""
            ComboBox2.Visible = False
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            ComboBox2.Visible = True
        End If
    End Sub


    Sub toexport()

        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument
        Dim tb1 As String = txt1.Text
        Dim tb2 As String = txt2.Text
        Dim tb3 As String = txt3.Text
        Dim tb4 As String = txt4.Text
        Dim selectedSource As String = cbx1.Text
        Dim selectedArea As String = ComboBox2.Text
        Dim str As Date = dtpStart.Value.ToString
        Dim stp As Date = dtpStop.Value.ToString
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3



        'Dim prodDes As String = tbxProdDes.Text

        If cb1.Text = "Receive Register" Then

            rd.Load(My.Application.Info.DirectoryPath + "\ireceive.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generatereceive(tb1, tb2, selectedSource, selectedArea, str, stp, rd)


        ElseIf cb1.Text = "Release Register" Then

            rd.Load(My.Application.Info.DirectoryPath + "\irelease.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generaterelease(tb1, tb2, selectedSource, selectedArea, str, stp, rd)

        ElseIf cb1.Text = "Build Inventory Register" Then

            rd.Load(My.Application.Info.DirectoryPath + "\iBuild Inventory Register.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateinventorybuildreg(tb1, tb2, selectedArea, str, stp, rd)

        ElseIf cb1.Text = "Item Register" Then
            rd.Load(My.Application.Info.DirectoryPath + "\Itemregister.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateItemRegister(tb1, str, stp, rd)


        ElseIf cb1.Text = "Transfer Register" Then
            rd.Load(My.Application.Info.DirectoryPath + "\IntraTransRegister.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateTransferRegister(tb1, tb2, selectedArea, str, stp, rd)


        ElseIf cb1.Text = "Count Register" Then
            rd.Load(My.Application.Info.DirectoryPath + "\icountregister.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateCountRegister(tb1, selectedArea, str, stp, rd)

        ElseIf cb1.Text = "Product Register" Then
            rd.Load(My.Application.Info.DirectoryPath + "\iProductRegister.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateProductRegister(tb1, tb2, selectedArea, rd)

        ElseIf cb1.Text = "Inventory List by Location" Then
            rd.Load(My.Application.Info.DirectoryPath + "\Inventory List by Location.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateInventorylistbylocation(tb1, tb2, selectedArea, tb3, tb4, rd)


        ElseIf cb1.Text = "Inventory List" Then
            rd.Load(My.Application.Info.DirectoryPath + "\Inventory List.rpt")
            rd.SetDatabaseLogon(RPuser, RPpass)
            q.generateInventorylist(tb1, tb2, selectedArea, rd)
        End If

        CrystalReportViewer1.ReportSource = rd
        CrystalReportViewer1.Refresh()
        CrystalReportViewer1.Zoom(83)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        toexport()
    End Sub

    Private Sub DtpStart_ValueChanged(sender As Object, e As EventArgs) Handles dtpStart.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
    End Sub

    Private Sub DtpStop_ValueChanged(sender As Object, e As EventArgs) Handles dtpStop.ValueChanged
        If dtpStart.Value > dtpStop.Value Then
            dtpStop.Value = dtpStart.Value
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Export_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        cb1.Items.Clear()
        txt1.Text = ""
        txt2.Text = ""
        txt3.Text = ""
        txt4.Text = ""
        RadioButton1.Checked = True
        RadioButton3.Checked = True
        CrystalReportViewer1.Refresh()


    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            cbx1.Visible = False
            RadioButton4.Checked = False
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            cbx1.Visible = True
            RadioButton3.Checked = False
        End If
    End Sub

    Private Sub Txt1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txt1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    e.Handled = True
        '    MessageBox.Show("
        '
        '
        '    ")
        'End If
    End Sub

    Private Sub Export_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class