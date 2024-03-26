Public Class confirmationFormAdjustInventory
    Private q As New qry
    Public approvedId As String = ""

    Private Sub confirmationFormAdjustInventory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        clearFields()
    End Sub

    Sub clearFields()
        tbxUser.Text = ""
        tbxPass.Text = ""
        tbxUser.Select()
    End Sub

    Sub closeThisForm()
        Home_Page.Enabled = True
        adjustInventory.Enabled = True
        Me.Dispose()
        adjustInventory.Focus()
        adjustInventory.Select()
    End Sub

    Private Sub confirmationFormAdjustInventory_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True
        adjustInventory.Enabled = True
    End Sub

    Private Sub confirmationFormAdjustInventory_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
                If (MessageBox.Show("Are you sure to cancel confimation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

                Else
                    closeThisForm()
                End If
            Else
                closeThisForm()
            End If
        End If
    End Sub

    Private Sub btnCanel_Click(sender As Object, e As EventArgs) Handles btnCanel.Click
        If Not String.IsNullOrWhiteSpace(tbxUser.Text) Or Not String.IsNullOrWhiteSpace(tbxPass.Text) Then
            If (MessageBox.Show("Are you sure to cancel confimation?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then

            Else
                closeThisForm()
            End If
        Else
            closeThisForm()
        End If
    End Sub

    Private Sub btnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        If String.IsNullOrWhiteSpace(tbxUser.Text) Or String.IsNullOrWhiteSpace(tbxPass.Text) Then
            MessageBox.Show("Fields with * are required.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        Else
            With adjustInventory
                '----------bushit ni na code gina earase niya ang gina search nmo pag search kay naa pero pag record na kay mawala na
                '.TextBox1.Text = ""
                '.TextBox2.Text = ""
                '.TextBox3.Text = ""
                .dgvProducts.ClearSelection()

            End With

            q.inventoryAdjustmentApproval(tbxUser.Text, tbxPass.Text, newHome.areaId)
        End If
    End Sub
End Class