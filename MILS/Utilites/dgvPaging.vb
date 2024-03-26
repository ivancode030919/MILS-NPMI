Public Class dgvPaging
    Public PageCount As Integer
    Private maxRec As Integer
    Private pageSize As Integer
    Public Property currentPage As Integer
    Private recNo As Integer

    Public Sub LoadPage(ByRef dt As DataTable, ByVal dgv As DataGridView, ByVal page As String)
        Dim i As Integer
        Dim startRec As Integer
        Dim endRec As Integer
        Dim dtTemp As DataTable
        'Dim dr As DataRow

        'Duplicate or clone the source table to create the temporary table.
        If Not dt.Rows.Count = 0 Then
            dtTemp = dt.Clone
        Else
            'MessageBox.Show("No data to show.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If


        If currentPage = PageCount Then
            endRec = maxRec
        Else
            endRec = pageSize * currentPage
        End If

        startRec = recNo

        'Copy the rows from the source table to fill the temporary table.
        For i = startRec To endRec - 1
            dtTemp.ImportRow(dt.Rows(i))
            recNo = recNo + 1
        Next

        dgv.DataSource = dtTemp
        'DisplayPageInfo(page)
    End Sub

    Public Sub DisplayPageInfo(ByVal lbl As Label)
        lbl.Text = "Page " & currentPage.ToString & "/ " & PageCount.ToString
    End Sub

    Private Function CheckFillButton() As Boolean
        'Check if the user clicks the "Fill Grid" button.
        If pageSize = 0 Then
            MessageBox.Show("Show data first!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            CheckFillButton = False
        Else
            CheckFillButton = True
        End If
    End Function

    Public Sub fillDgv(ByVal pages As String, ByVal dt As DataTable, Optional ByVal i As Integer = 1)
        'Set the start and max records. 
        pageSize = pages
        maxRec = dt.Rows.Count
        PageCount = maxRec \ pageSize

        ' Adjust the page number if the last page contains a partial page.
        If (maxRec Mod pageSize) > 0 Then
            PageCount = PageCount + 1
        End If

        'Initial seeings
        currentPage = i
        recNo = 0

        'LoadPage()
    End Sub

    Public Sub nextpage(ByRef dt As DataTable, ByVal dgv As DataGridView, ByVal page As String)
        'If the user did not click the "Fill Grid" button then Return
        If Not CheckFillButton() Then Return

        'Check if the user clicked the "Fill Grid" button.
        If pageSize = 0 Then
            'MessageBox.Show("Set the Page Size, and then click the ""Fill Grid"" button!")
            Return
        End If

        currentPage = currentPage + 1

        If currentPage > PageCount Then
            currentPage = PageCount

            'Check if you are already at the last page.
            If recNo = maxRec Then
                'MessageBox.Show("You are at the Last Page!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Return
            End If
        End If
        LoadPage(dt, dgv, page)
    End Sub

    Public Sub previouspage(ByRef dt As DataTable, ByVal dgv As DataGridView, ByVal page As String)
        If Not CheckFillButton() Then Return

        If currentPage = PageCount Then
            recNo = pageSize * (currentPage - 2)
        End If

        currentPage = currentPage - 1

        'Check if you are already at the first page.
        If currentPage < 1 Then
            'MessageBox.Show("You are at the First Page!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            currentPage = 1
            Return
        Else
            recNo = pageSize * (currentPage - 1)
        End If
        LoadPage(dt, dgv, page)
    End Sub

    Public Sub tofirstpage(ByRef dt As DataTable, ByVal dgv As DataGridView, ByVal page As String)
        If Not CheckFillButton() Then Return

        ' Check if you are already at the first page.
        If currentPage = 1 Then
            'MessageBox.Show("You are at the First Page!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        currentPage = 1
        recNo = 0

        LoadPage(dt, dgv, page)
    End Sub

    Public Sub tolastpage(ByRef dt As DataTable, ByVal dgv As DataGridView, ByVal page As String)
        If Not CheckFillButton() Then Return

        ' Check if you are already at the last page.
        If recNo = maxRec Then
            'MessageBox.Show("You are at the Last Page!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return
        End If

        currentPage = PageCount

        recNo = pageSize * (currentPage - 1)

        LoadPage(dt, dgv, page)
    End Sub
End Class
