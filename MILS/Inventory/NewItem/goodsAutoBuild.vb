Public Class goodsAutoBuild
    Private q As New qry
    Private table As New DataTable("Table")
    Public apprvdBy As String = ""

    Private Sub goodsAutoBuild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtpDateEnc.MaxDate = DateTime.Now
        newDGVFormat()
        dtpDateEnc.Select()
    End Sub

    Sub newDGVFormat()
        'With table
        '    .Columns.Add("PRODUCT NUMBER", Type.GetType("System.String"))  '0
        '    .Columns.Add("DESCRIPTION", Type.GetType("System.String"))     '1
        'End With
        With dgvGoods

            .Columns.Add("0", "PRODUCT NUMBER")
            .Columns.Add("1", "DESCRIPTION")

            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = True
            .EnableHeadersVisualStyles = False
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            .DataSource = table
            .Columns(0).Width = 170
            .Columns(1).Width = 547
            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            For index As Integer = 0 To .ColumnCount - 1
                .Columns(index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            Next
        End With
    End Sub

    Private Sub dgvGoods_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles dgvGoods.RowValidating
        Dim row As DataGridViewRow = dgvGoods.Rows(e.RowIndex)
        Dim isEmpty As Boolean = False
        Dim isRepeatingId As Boolean = False
        Dim alreadyInDBID As Boolean = False
        Dim idDGV As String = ""
        With dgvGoods
            If .IsCurrentRowDirty Then
                idDGV = row.Cells(0).Value.ToString
                For Each cell As DataGridViewCell In row.Cells 'is null
                    If String.IsNullOrWhiteSpace(cell.Value.ToString) Then
                        'e.Cancel = True
                        isEmpty = True
                    End If
                Next

                For i As Integer = 0 To Me.dgvGoods.RowCount - 1
                    For j As Integer = 0 To Me.dgvGoods.RowCount - 1
                        If i <> j Then
                            Dim cellValueI As Object = dgvGoods.Rows(i).Cells(0).Value
                            Dim cellValueJ As Object = dgvGoods.Rows(j).Cells(0).Value

                            If cellValueI IsNot Nothing AndAlso cellValueJ IsNot Nothing AndAlso cellValueI.Equals(cellValueJ) Then
                                dgvGoods.Rows(i).DefaultCellStyle.BackColor = Color.Red
                                isRepeatingId = True
                            Else
                                dgvGoods.Rows(i).DefaultCellStyle.BackColor = Color.White
                            End If
                        End If
                    Next
                Next

                Try
                    alreadyInDBID = q.validateIDinGoodsAutoBuild(idDGV)
                Catch ex As Exception
                    MessageBox.Show("Please contact Project Team",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

                If isEmpty = True Then
                    e.Cancel = True
                    MessageBox.Show("Fields with asterisk are required.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ElseIf isRepeatingId = True Then
                    e.Cancel = True
                    MessageBox.Show("Duplicate Product Number are not valid.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                ElseIf alreadyInDBID = True Then
                    e.Cancel = True
                    MessageBox.Show("Product Number already exists.",
                                  "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End If
        End With
    End Sub

    Private Sub goodsAutoBuild_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If dgvGoods.Rows.Count > 0 Then
            If (MessageBox.Show("Are you sure you want to cancel transaction?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                e.Cancel = True
            Else
                e.Cancel = False
                dgvGoods.Columns.Clear()
            End If
        End If

    End Sub

    Private Sub validateDatas()
        If dgvGoods.Rows.Count > 1 Then
            Home_Page.Enabled = False
            Me.Enabled = False
            With goodsAutoBuildConfirmation
                .Show()
                .Focus()
                .Select()
            End With
        Else
            MessageBox.Show("Input valid data.",
                                "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
    End Sub

    Sub addDatas()
        q.addNewGoodsAutoBuild(tbxRemarks.Text, dtpDateEnc.Value, newHome.userId, apprvdBy)
    End Sub


    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnRecord.Click
        If btnRecord.Text = "Record" Then

            If (MessageBox.Show("Do You Want To Record?", "Confirmation", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else

                validateDatas()

            End If

        ElseIf btnRecord.Text = "Add New" Then
            With Me

                .dgvGoods.Enabled = True
                .dgvGoods.AllowUserToAddRows = True
                .table.Clear()
                .tbxEntry.Text = ""
                .tbxRemarks.Text = ""
                .lblErr.Visible = False
                .btnRecord.Text = "Record"

            End With
        End If
    End Sub

    Private Sub dgvGoods_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvGoods.KeyDown
        If dgvGoods.Rows.Count > 1 Then
            If e.KeyCode = Keys.Delete Then
                If dgvGoods.SelectedRows.Count > 0 Then
                    'you may want to add a confirmation message, and if the user confirms delete
                    dgvGoods.Rows.Remove(dgvGoods.SelectedRows(0))
                Else
                    MessageBox.Show("Select valid row to delete.")
                End If
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub GoodsAutoBuild_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class