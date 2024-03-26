Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Drawing.Image
Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms


Public Class UpdateDes

    Private q As New qry
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If String.IsNullOrWhiteSpace(TextBox1.Text) Then
            Label3.Visible = True
            Label3.Text = "Please Select Product !!!"

        Else
            Dim update As New UpdateConfirmation()
            Me.Enabled = False
            update.ShowDialog()
            TextBox1.Text = update.TextBox1.Text
            destext.Text = update.TextBox2.Text
        End If
    End Sub

    Private Sub UpdateDes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadlv()
        newLVformat1()
        Label3.Visible = False
    End Sub


    Sub loadlv()
        q.updateloadGoodsToEdit(ListView1, TextBox1.Text, destext.Text)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        loadlv()
    End Sub

    Private Sub Destext_TextChanged(sender As Object, e As EventArgs) Handles destext.TextChanged
        loadlv()
    End Sub


    Sub newLVformat1()
        With ListView1
            .Columns.Add("Product Number", 130, HorizontalAlignment.Center)
            .Columns.Add("Descrition", 500, HorizontalAlignment.Left)
            .Columns(0).TextAlign = HorizontalAlignment.Center
        End With
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        destext.Text = ""
        Label3.Visible = False
    End Sub

    Private Sub textBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        'If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
        '    ' If the entered key is not a digit or a control key (e.g., Backspace)
        '    e.Handled = True ' Ignore the keypress event
        'End If
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems.Count > 0 Then
            Dim selectedItem As ListViewItem = ListView1.SelectedItems(0)
            Dim itemID As String = selectedItem.SubItems(0).Text ' Assuming ID is in the first subitem column
            Dim itemdes As String = selectedItem.SubItems(1).Text ' Assuming ID is in the first subitem column
            TextBox1.Text = itemID
            destext.Text = itemdes


            Label3.Text = "*"
            Label3.Visible = False
        Else
            TextBox1.Text = ""
            destext.Text = ""
        End If
    End Sub

    Private Sub UpdateDes_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ListView1.Clear()
    End Sub

    Private Sub UpdateDes_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    'Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
    '    If ListView1.SelectedItems.Count > 0 Then
    '        TextBox1.Text = ListView1.SelectedItems(0).SubItems(0).Text
    '        destext.Text = ListView1.SelectedItems(0).SubItems(1).Text
    '    End If
    'End Sub

    'Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
    '    Dim selectedItem As ListViewItem = ListView1.SelectedItems(0)
    '    TextBox1.Text = 
    'End Sub
End Class