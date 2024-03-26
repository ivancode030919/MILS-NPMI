<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class confirmationFormAdjustInventory
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
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblErrPASS
        '
        Me.lblErrPASS.AutoSize = True
        Me.lblErrPASS.ForeColor = System.Drawing.Color.Red
        Me.lblErrPASS.Location = New System.Drawing.Point(82, 33)
        Me.lblErrPASS.Name = "lblErrPASS"
        Me.lblErrPASS.Size = New System.Drawing.Size(13, 13)
        Me.lblErrPASS.TabIndex = 16
        Me.lblErrPASS.Text = "*"
        '
        'lblErrUN
        '
        Me.lblErrUN.AutoSize = True
        Me.lblErrUN.ForeColor = System.Drawing.Color.Red
        Me.lblErrUN.Location = New System.Drawing.Point(82, 5)
        Me.lblErrUN.Name = "lblErrUN"
        Me.lblErrUN.Size = New System.Drawing.Size(13, 13)
        Me.lblErrUN.TabIndex = 15
        Me.lblErrUN.Text = "*"
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(11, 65)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(13, 13)
        Me.lblError.TabIndex = 14
        Me.lblError.Text = "*"
        Me.lblError.Visible = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(57, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Password:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(59, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Username:"
        '
        'btnLogin
        '
        Me.btnLogin.Location = New System.Drawing.Point(169, 77)
        Me.btnLogin.Name = "btnLogin"
        Me.btnLogin.Size = New System.Drawing.Size(75, 32)
        Me.btnLogin.TabIndex = 11
        Me.btnLogin.Text = "Confirm"
        Me.btnLogin.UseVisualStyleBackColor = True
        '
        'btnCanel
        '
        Me.btnCanel.Location = New System.Drawing.Point(250, 77)
        Me.btnCanel.Name = "btnCanel"
        Me.btnCanel.Size = New System.Drawing.Size(75, 32)
        Me.btnCanel.TabIndex = 13
        Me.btnCanel.Text = "Cancel"
        Me.btnCanel.UseVisualStyleBackColor = True
        '
        'tbxPass
        '
        Me.tbxPass.Location = New System.Drawing.Point(104, 33)
        Me.tbxPass.Name = "tbxPass"
        Me.tbxPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbxPass.Size = New System.Drawing.Size(221, 21)
        Me.tbxPass.TabIndex = 9
        '
        'tbxUser
        '
        Me.tbxUser.Location = New System.Drawing.Point(104, 5)
        Me.tbxUser.Name = "tbxUser"
        Me.tbxUser.Size = New System.Drawing.Size(221, 21)
        Me.tbxUser.TabIndex = 8
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.tbxUser)
        Me.PanelControl1.Controls.Add(Me.lblErrPASS)
        Me.PanelControl1.Controls.Add(Me.tbxPass)
        Me.PanelControl1.Controls.Add(Me.lblErrUN)
        Me.PanelControl1.Controls.Add(Me.btnCanel)
        Me.PanelControl1.Controls.Add(Me.lblError)
        Me.PanelControl1.Controls.Add(Me.btnLogin)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(335, 117)
        Me.PanelControl1.TabIndex = 17
        '
        'confirmationFormAdjustInventory
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(335, 117)
        Me.Controls.Add(Me.PanelControl1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "confirmationFormAdjustInventory"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Confirmation"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.ResumeLayout(False)

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
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
