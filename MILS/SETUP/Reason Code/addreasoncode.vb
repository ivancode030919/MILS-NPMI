
Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms

Public Class addreasoncode
    Private q As New qry
    Private Sub Addreasoncode_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbvalue()
    End Sub

    Sub cbvalue()
        If Not cb1.Items.Contains("New Inventory") Then
            cb1.Items.Add("New Inventory")
        End If

        If Not cb1.Items.Contains("Inventory Adjustment") Then
            cb1.Items.Add("New Count")
        End If

        If Not cb1.Items.Contains("Receive") Then
            cb1.Items.Add("Receive Goods")
        End If

        If Not cb1.Items.Contains("Release") Then
            cb1.Items.Add("Release Goods")
        End If

        If Not cb1.Items.Contains("Internal Warehouse Transfer") Then
            cb1.Items.Add("Transfer Goods")
        End If


    End Sub


    Private Sub addreasoncode()


        If txt1.Text = "" Then
            MessageBox.Show("Please Put the Reason Code")
        ElseIf txt2.Text = "" Then
            MessageBox.Show("Please Put the Reason Code Description")
        Else

            q.addreasoncode(txt1.Text, txt2.Text, cb1.Text)

        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        addreasoncode()
    End Sub

    Private Sub Addreasoncode_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        txt1.Text = ""
        txt2.Text = ""
        cb1.SelectedItem = -1
        cb1.Items.Clear()
    End Sub
End Class