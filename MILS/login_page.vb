
'developers::::::::::::::::::::
'nag start -JC
'preceedor - junie cool
'april 2023 - ivan bajenting

Public Class login_page
    Private q As New qry
    Private Sub login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        tbxUser.Select()
        tbxUser.Focus()

        Timer1.Enabled = True

    End Sub

    Sub validateUser()
        If String.IsNullOrWhiteSpace(tbxUser.Text) And String.IsNullOrWhiteSpace(tbxPass.Text) Then

            MessageBox.Show("Username and Password is required", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else

            'q.logInUser(tbxUsernam.Text, tbxPassword.Text)
            q.logInUser(tbxUser.Text, tbxPass.Text)
        End If
    End Sub



    Private Sub btnCanel_Click(sender As Object, e As EventArgs) Handles btnCanel.Click
        Dim btnClose As Boolean = True
        If (MessageBox.Show("Are you sure you want to close this application?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
            btnClose = True
        Else
            btnClose = False
            Application.Exit()
        End If
    End Sub

    Private Sub login_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim btnClose As Boolean = True
        If e.KeyCode = Keys.Escape Then
            If (MessageBox.Show("Are you sure you want to close this application?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                btnClose = True
            Else
                btnClose = False
                Application.Exit()
            End If
        ElseIf e.KeyCode = Keys.Enter Then
            validateUser()
        End If
    End Sub

    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles btnLogin.Click
        validateUser()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

    End Sub
End Class