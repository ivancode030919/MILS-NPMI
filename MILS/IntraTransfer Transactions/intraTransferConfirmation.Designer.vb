<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class intraTransferConfirmation
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lblErrPASS = New System.Windows.Forms.Label()
        Me.lblErrUN = New System.Windows.Forms.Label()
        Me.lblError = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLogin = New System.Windows.Forms.Button()
        Me.btnCanel = New System.Windows.Forms.Button()
        Me.tbxPass = New System.Windows.Forms.TextBox()
        Me.tbxUser = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblErrPASS
        '
        Me.lblErrPASS.AutoSize = True
        Me.lblErrPASS.ForeColor = System.Drawing.Color.Red
        Me.lblErrPASS.Location = New System.Drawing.Point(286, 53)
        Me.lblErrPASS.Name = "lblErrPASS"
        Me.lblErrPASS.Size = New System.Drawing.Size(16, 17)
        Me.lblErrPASS.TabIndex = 25
        Me.lblErrPASS.Text = "*"
        '
        'lblErrUN
        '
        Me.lblErrUN.AutoSize = True
        Me.lblErrUN.ForeColor = System.Drawing.Color.Red
        Me.lblErrUN.Location = New System.Drawing.Point(286, 19)
        Me.lblErrUN.Name = "lblErrUN"
        Me.lblErrUN.Size = New System.Drawing.Size(16, 17)
        Me.lblErrUN.TabIndex = 24
        Me.lblErrUN.Text = "*"
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(103, 77)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(16, 17)
        Me.lblError.TabIndex = 23
        Me.lblError.Text = "*"
        Me.lblError.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(14, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 18)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Password :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 18)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Username :"
        '
        'btnLogin
        '
        Me.btnLogin.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLogin.Location = New System.Drawing.Point(207, 97)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(75, 32)
        Me.btnLogin.TabIndex = 20
        Me.btnLogin.Text = "Confirm"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'btnCanel
        '
        Me.btnCanel.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCanel.Location = New System.Drawing.Point(126, 97)
        Me.btnCanel.Name = "btnCanel"
        Me.btnCanel.Size = New System.Drawing.Size(75, 32)
        Me.btnCanel.TabIndex = 22
        Me.btnCanel.Text = "Cancel"
        Me.btnCanel.UseVisualStyleBackColor = True
        '
        'tbxPass
        '
        Me.tbxPass.Dock = System.Windows.Forms.DockStyle.Top
        Me.tbxPass.Location = New System.Drawing.Point(0, 0)
        Me.tbxPass.Name = "tbxPass"
        Me.tbxPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbxPass.Size = New System.Drawing.Size(174, 22)
        Me.tbxPass.TabIndex = 18
        '
        'tbxUser
        '
        Me.tbxUser.Dock = System.Windows.Forms.DockStyle.Top
        Me.tbxUser.Location = New System.Drawing.Point(0, 0)
        Me.tbxUser.Name = "tbxUser"
        Me.tbxUser.Size = New System.Drawing.Size(174, 22)
        Me.tbxUser.TabIndex = 17
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.tbxPass)
        Me.Panel1.Location = New System.Drawing.Point(102, 45)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(178, 25)
        Me.Panel1.TabIndex = 26
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.tbxUser)
        Me.Panel2.Location = New System.Drawing.Point(102, 16)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(178, 26)
        Me.Panel2.TabIndex = 27
        '
        'intraTransferConfirmation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(310, 153)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblErrPASS)
        Me.Controls.Add(Me.lblErrUN)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnLogin)
        Me.Controls.Add(Me.btnCanel)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "intraTransferConfirmation"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Confirmation"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblErrPASS As Label
    Friend WithEvents lblErrUN As Label
    Friend WithEvents lblError As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents btnLogin As Button
    Friend WithEvents btnCanel As Button
    Friend WithEvents tbxPass As TextBox
    Friend WithEvents tbxUser As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
End Class
