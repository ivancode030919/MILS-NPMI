<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class usersDetails
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
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.rdInactive = New System.Windows.Forms.RadioButton()
        Me.rdActive = New System.Windows.Forms.RadioButton()
        Me.lblErrArea = New System.Windows.Forms.Label()
        Me.lblErrRole = New System.Windows.Forms.Label()
        Me.lblErrLN = New System.Windows.Forms.Label()
        Me.lblErrFN = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbxFN = New System.Windows.Forms.TextBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxRole = New System.Windows.Forms.ComboBox()
        Me.tbxMN = New System.Windows.Forms.TextBox()
        Me.tbxLN = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnUpdate = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbxPass = New System.Windows.Forms.TextBox()
        Me.tbxUserName = New System.Windows.Forms.TextBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(584, 299)
        Me.TabControl1.TabIndex = 1
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label10)
        Me.TabPage1.Controls.Add(Me.rdInactive)
        Me.TabPage1.Controls.Add(Me.rdActive)
        Me.TabPage1.Controls.Add(Me.lblErrArea)
        Me.TabPage1.Controls.Add(Me.lblErrRole)
        Me.TabPage1.Controls.Add(Me.lblErrLN)
        Me.TabPage1.Controls.Add(Me.lblErrFN)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.tbxFN)
        Me.TabPage1.Controls.Add(Me.cbxArea)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.cbxRole)
        Me.TabPage1.Controls.Add(Me.tbxMN)
        Me.TabPage1.Controls.Add(Me.tbxLN)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Location = New System.Drawing.Point(4, 26)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(576, 269)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Details"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(67, 192)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(46, 17)
        Me.Label10.TabIndex = 26
        Me.Label10.Text = "Status:"
        '
        'rdInactive
        '
        Me.rdInactive.AutoSize = True
        Me.rdInactive.Location = New System.Drawing.Point(174, 217)
        Me.rdInactive.Name = "rdInactive"
        Me.rdInactive.Size = New System.Drawing.Size(69, 21)
        Me.rdInactive.TabIndex = 25
        Me.rdInactive.Text = "Inactive"
        Me.rdInactive.UseVisualStyleBackColor = True
        '
        'rdActive
        '
        Me.rdActive.AutoSize = True
        Me.rdActive.Checked = True
        Me.rdActive.Location = New System.Drawing.Point(174, 190)
        Me.rdActive.Name = "rdActive"
        Me.rdActive.Size = New System.Drawing.Size(59, 21)
        Me.rdActive.TabIndex = 24
        Me.rdActive.TabStop = True
        Me.rdActive.Text = "Active"
        Me.rdActive.UseVisualStyleBackColor = True
        '
        'lblErrArea
        '
        Me.lblErrArea.AutoSize = True
        Me.lblErrArea.ForeColor = System.Drawing.Color.Red
        Me.lblErrArea.Location = New System.Drawing.Point(152, 159)
        Me.lblErrArea.Name = "lblErrArea"
        Me.lblErrArea.Size = New System.Drawing.Size(16, 17)
        Me.lblErrArea.TabIndex = 23
        Me.lblErrArea.Text = "*"
        '
        'lblErrRole
        '
        Me.lblErrRole.AutoSize = True
        Me.lblErrRole.ForeColor = System.Drawing.Color.Red
        Me.lblErrRole.Location = New System.Drawing.Point(152, 128)
        Me.lblErrRole.Name = "lblErrRole"
        Me.lblErrRole.Size = New System.Drawing.Size(16, 17)
        Me.lblErrRole.TabIndex = 22
        Me.lblErrRole.Text = "*"
        '
        'lblErrLN
        '
        Me.lblErrLN.AutoSize = True
        Me.lblErrLN.ForeColor = System.Drawing.Color.Red
        Me.lblErrLN.Location = New System.Drawing.Point(152, 99)
        Me.lblErrLN.Name = "lblErrLN"
        Me.lblErrLN.Size = New System.Drawing.Size(16, 17)
        Me.lblErrLN.TabIndex = 21
        Me.lblErrLN.Text = "*"
        '
        'lblErrFN
        '
        Me.lblErrFN.AutoSize = True
        Me.lblErrFN.ForeColor = System.Drawing.Color.Red
        Me.lblErrFN.Location = New System.Drawing.Point(152, 43)
        Me.lblErrFN.Name = "lblErrFN"
        Me.lblErrFN.Size = New System.Drawing.Size(16, 17)
        Me.lblErrFN.TabIndex = 20
        Me.lblErrFN.Text = "*"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(67, 162)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(36, 17)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Area:"
        '
        'tbxFN
        '
        Me.tbxFN.Location = New System.Drawing.Point(174, 43)
        Me.tbxFN.Name = "tbxFN"
        Me.tbxFN.Size = New System.Drawing.Size(300, 22)
        Me.tbxFN.TabIndex = 11
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(174, 159)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(300, 25)
        Me.cbxArea.TabIndex = 18
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(67, 46)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "First Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(67, 131)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(36, 17)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Role:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(67, 74)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Middle Name:"
        '
        'cbxRole
        '
        Me.cbxRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxRole.FormattingEnabled = True
        Me.cbxRole.Location = New System.Drawing.Point(174, 128)
        Me.cbxRole.Name = "cbxRole"
        Me.cbxRole.Size = New System.Drawing.Size(300, 25)
        Me.cbxRole.TabIndex = 16
        '
        'tbxMN
        '
        Me.tbxMN.Location = New System.Drawing.Point(174, 71)
        Me.tbxMN.Name = "tbxMN"
        Me.tbxMN.Size = New System.Drawing.Size(300, 22)
        Me.tbxMN.TabIndex = 13
        '
        'tbxLN
        '
        Me.tbxLN.Location = New System.Drawing.Point(174, 99)
        Me.tbxLN.Name = "tbxLN"
        Me.tbxLN.Size = New System.Drawing.Size(300, 22)
        Me.tbxLN.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(67, 102)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(71, 17)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Last Name:"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label9)
        Me.TabPage2.Controls.Add(Me.Label8)
        Me.TabPage2.Controls.Add(Me.btnCancel)
        Me.TabPage2.Controls.Add(Me.btnUpdate)
        Me.TabPage2.Controls.Add(Me.Label7)
        Me.TabPage2.Controls.Add(Me.Label6)
        Me.TabPage2.Controls.Add(Me.tbxPass)
        Me.TabPage2.Controls.Add(Me.tbxUserName)
        Me.TabPage2.Location = New System.Drawing.Point(4, 26)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(576, 269)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Account"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(148, 109)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(16, 17)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "*"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Red
        Me.Label8.Location = New System.Drawing.Point(148, 81)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(16, 17)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "*"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(395, 210)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnUpdate
        '
        Me.btnUpdate.Location = New System.Drawing.Point(314, 210)
        Me.btnUpdate.Name = "btnUpdate"
        Me.btnUpdate.Size = New System.Drawing.Size(75, 25)
        Me.btnUpdate.TabIndex = 4
        Me.btnUpdate.Text = "Update"
        Me.btnUpdate.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(63, 112)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 17)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Password:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(63, 84)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(68, 17)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Username:"
        '
        'tbxPass
        '
        Me.tbxPass.Location = New System.Drawing.Point(170, 109)
        Me.tbxPass.Name = "tbxPass"
        Me.tbxPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbxPass.Size = New System.Drawing.Size(300, 22)
        Me.tbxPass.TabIndex = 1
        '
        'tbxUserName
        '
        Me.tbxUserName.Location = New System.Drawing.Point(170, 81)
        Me.tbxUserName.Name = "tbxUserName"
        Me.tbxUserName.Size = New System.Drawing.Size(300, 22)
        Me.tbxUserName.TabIndex = 0
        '
        'usersDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(584, 299)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "usersDetails"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Details"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents lblErrArea As Label
    Friend WithEvents lblErrRole As Label
    Friend WithEvents lblErrLN As Label
    Friend WithEvents lblErrFN As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents tbxFN As TextBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cbxRole As ComboBox
    Friend WithEvents tbxMN As TextBox
    Friend WithEvents tbxLN As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnUpdate As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents tbxPass As TextBox
    Friend WithEvents tbxUserName As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents rdInactive As RadioButton
    Friend WithEvents rdActive As RadioButton
End Class
