Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class itemListSmmary
    Private q As New qry
    Public RPuser3 As String
    Public RPpass3 As String
    Private Sub itemListSmmary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadCBXProducts()
        q.reportlog()
        'q.loadAreasInReport(clbAreas)
    End Sub

    Private Sub loadCBXProducts()
        q.loadProductListforItemListSummary(cbxProdDes, tbxProdId, tbxProdDes)
    End Sub
    Private Sub generateReport()
        Dim isChecked As Boolean = False
        Dim selectedPDNO As String = ""
        Dim isCheckedArea As Boolean = False
        Dim selectedArea As String = ""
        Dim RPuser As String = RPuser3
        Dim RPpass As String = RPpass3

        For i As Integer = 0 To cbxProdDes.Items.Count - 1
            If cbxProdDes.GetItemChecked(i) = True Then
                isChecked = True
                Exit For
            Else
            End If
        Next

        For j As Integer = 0 To clbAreas.Items.Count - 1
            If clbAreas.GetItemChecked(j) = True Then
                isCheckedArea = True
                Exit For
            Else
            End If
        Next

        If isChecked = True Then
            For Each chk As String In cbxProdDes.CheckedItems
                Dim firstseparator As String = "("
                Dim firstseparatorIndex = chk.IndexOf(firstseparator) + 1

                Dim secondSeparator As String = ")"
                Dim secondSeparatorIndex = chk.IndexOf(secondSeparator) - 1

                selectedPDNO &= "'" + chk.Substring(firstseparatorIndex, secondSeparatorIndex) & "', "
            Next
            selectedPDNO = selectedPDNO.Remove(selectedPDNO.Length - 2, 2)
        ElseIf isChecked = False Then
            For Each chk As String In cbxProdDes.Items
                Dim firstseparator As String = "("
                Dim firstseparatorIndex = chk.IndexOf(firstseparator) + 1

                Dim secondSeparator As String = ")"
                Dim secondSeparatorIndex = chk.IndexOf(secondSeparator) - 1

                selectedPDNO &= "'" + chk.Substring(firstseparatorIndex, secondSeparatorIndex) & "', "
            Next
            selectedPDNO = selectedPDNO.Remove(selectedPDNO.Length - 2, 2)
        End If

        'Get the value of selected area
        If isCheckedArea = True Then
            For Each chk As String In clbAreas.CheckedItems
                selectedArea &= "'" + chk & "', "
            Next
            selectedArea = selectedArea.Remove(selectedArea.Length - 2, 2)
        ElseIf isCheckedArea = False Then
            For Each chk As String In clbAreas.Items
                selectedArea &= "'" + chk + "', "
            Next
            selectedArea = selectedArea.Remove(selectedArea.Length - 2, 2)
        End If
        'MsgBox(selectedArea + " " + selectedPDNO)
        Dim crTableLogoninfos As New TableLogOnInfos
        Dim crTableLogoninfo As New TableLogOnInfo
        Dim crConnectionInfo As New ConnectionInfo
        Dim rd As New ReportDocument

        rd.Load(My.Application.Info.DirectoryPath + "\itemListSummary.rpt")
        'rd.SetDatabaseLogon("sa", "p@ssw0rd")

        q.generateItemListSummary(selectedPDNO, selectedArea, rd)

        crvIGLR.ReportSource = rd
        crvIGLR.Refresh()
        crvIGLR.Zoom(107)
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Try
            generateReport()
        Catch ex As Exception
            MessageBox.Show("Inventory is on zero balance.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub tbxProdId_TextChanged(sender As Object, e As EventArgs) Handles tbxProdId.TextChanged
        loadCBXProducts()
    End Sub

    Private Sub tbxProdDes_TextChanged(sender As Object, e As EventArgs) Handles tbxProdDes.TextChanged
        loadCBXProducts()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        tbxProdId.Text = ""
        tbxProdDes.Text = ""
        tbxProdId.Select()

        For i As Integer = 0 To cbxProdDes.Items.Count - 1
            cbxProdDes.SetItemChecked(i, False)
        Next

        For i As Integer = 0 To clbAreas.Items.Count - 1
            clbAreas.SetItemChecked(i, False)
        Next

        cbxProdDes.SelectedItems.Clear()
        clbAreas.SelectedItems.Clear()
    End Sub

    Private Sub crvIGLR_Load(sender As Object, e As EventArgs) Handles crvIGLR.Load
        Size = New Size(928, 641)
    End Sub

    Private Sub ItemListSmmary_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Home_Page.MenuStrip1.Enabled = True
    End Sub
End Class