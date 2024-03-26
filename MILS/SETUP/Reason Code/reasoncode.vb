
Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms

Public Class reasoncode

    Private q As New qry
    Private Sub Reasoncode_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        lvcolumns()
        view()
        cb1.Visible = False
        RadioButton1.Checked = True
    End Sub


    Sub cbvalue()
        'cb1.Items.AddRange({"New Invemtory", "Inventory Adjustment"})

        cb1.Items.Add("New Inventory")

        cb1.Items.Add("New Count")
        cb1.Items.Add("Receive Goods")
        cb1.Items.Add("Release Goods")
        cb1.Items.Add("Transfer Goods")
    End Sub

    Sub lvcolumns()
        With lvrc
            .Columns.Add("ID", 0)
            .Columns.Add("REMARKS CODE", 150)
            .Columns.Add("DESCRIPTION", 200)
            .Columns.Add("MODULE", 150)
        End With
    End Sub

    Public Sub view()
        q.viewreasoncode(lvrc, txt1.Text, txt2.Text, cb1.Text)

    End Sub


    Private Sub Txt1_TextChanged(sender As Object, e As EventArgs) Handles txt1.TextChanged
        view()
    End Sub

    Private Sub Txt2_TextChanged(sender As Object, e As EventArgs) Handles txt2.TextChanged
        view()
    End Sub

    Private Sub Cb1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cb1.SelectedIndexChanged
        view()
    End Sub

    Private Sub Lvrc_Click(sender As Object, e As EventArgs) Handles lvrc.Click


        With lvrc.SelectedItems.Item(0)
            txt1.Text = .SubItems(1).Text
            txt2.Text = .SubItems(2).Text
            TextBox2.Text = .SubItems(0).Text
        End With
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Lvrc_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvrc.SelectedIndexChanged

    End Sub

    Private Sub reasoncode_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        lvrc.Clear()
        cb1.SelectedItem = -1
        cb1.Items.Clear()
    End Sub

    Private Sub BtnDetails_Click(sender As Object, e As EventArgs) Handles btnDetails.Click
        updatereasoncode.Show()
    End Sub

    Private Sub BtnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        addreasoncode.ShowDialog()
        RadioButton1.Checked = True
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        txt1.Text = ""
        txt2.Text = ""
        TextBox2.Text = ""
        RadioButton1.Checked = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
            cb1.SelectedIndex = -1
            cb1.Visible = False
            cb1.Items.Clear()
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            cb1.Visible = True
            cbvalue()
        End If
    End Sub

    Private Sub Reasoncode_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class