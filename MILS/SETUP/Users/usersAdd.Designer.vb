<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class usersAdd
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
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbxPass = New System.Windows.Forms.TextBox()
        Me.tbxUserName = New System.Windows.Forms.TextBox()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
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
        Me.TabControl1.Size = New System.Drawing.Size(445, 219)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.PanelControl1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 26)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(1157, 605)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Details"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'lblErrArea
        '
        Me.lblErrArea.AutoSize = True
        Me.lblErrArea.ForeColor = System.Drawing.Color.Red
        Me.lblErrArea.Location = New System.Drawing.Point(97, 124)
        Me.lblErrArea.Name = "lblErrArea"
        Me.lblErrArea.Size = New System.Drawing.Size(13, 13)
        Me.lblErrArea.TabIndex = 23
        Me.lblErrArea.Text = "*"
        '
        'lblErrRole
        '
        Me.lblErrRole.AutoSize = True
        Me.lblErrRole.ForeColor = System.Drawing.Color.Red
        Me.lblErrRole.Location = New System.Drawing.Point(97, 93)
        Me.lblErrRole.Name = "lblErrRole"
        Me.lblErrRole.Size = New System.Drawing.Size(13, 13)
        Me.lblErrRole.TabIndex = 22
        Me.lblErrRole.Text = "*"
        '
        'lblErrLN
        '
        Me.lblErrLN.AutoSize = True
        Me.lblErrLN.ForeColor = System.Drawing.Color.Red
        Me.lblErrLN.Location = New System.Drawing.Point(97, 64)
        Me.lblErrLN.Name = "lblErrLN"
        Me.lblErrLN.Size = New System.Drawing.Size(13, 13)
        Me.lblErrLN.TabIndex = 21
        Me.lblErrLN.Text = "*"
        '
        'lblErrFN
        '
        Me.lblErrFN.AutoSize = True
        Me.lblErrFN.ForeColor = System.Drawing.Color.Red
        Me.lblErrFN.Location = New System.Drawing.Point(97, 8)
        Me.lblErrFN.Name = "lblErrFN"
        Me.lblErrFN.Size = New System.Drawing.Size(13, 13)
        Me.lblErrFN.TabIndex = 20
        Me.lblErrFN.Text = "*"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(12, 127)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(34, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Area:"
        '
        'tbxFN
        '
        Me.tbxFN.Location = New System.Drawing.Point(119, 8)
        Me.tbxFN.Name = "tbxFN"
        Me.tbxFN.Size = New System.Drawing.Size(300, 21)
        Me.tbxFN.TabIndex = 11
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(119, 124)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(300, 21)
        Me.cbxArea.TabIndex = 18
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(62, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "First Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(12, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(32, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "Role:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Middle Name:"
        '
        'cbxRole
        '
        Me.cbxRole.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxRole.FormattingEnabled = True
        Me.cbxRole.Location = New System.Drawing.Point(119, 93)
        Me.cbxRole.Name = "cbxRole"
        Me.cbxRole.Size = New System.Drawing.Size(300, 21)
        Me.cbxRole.TabIndex = 16
        '
        'tbxMN
        '
        Me.tbxMN.Location = New System.Drawing.Point(119, 36)
        Me.tbxMN.Name = "tbxMN"
        Me.tbxMN.Size = New System.Drawing.Size(300, 21)
        Me.tbxMN.TabIndex = 13
        '
        'tbxLN
        '
        Me.tbxLN.Location = New System.Drawing.Point(119, 64)
        Me.tbxLN.Name = "tbxLN"
        Me.tbxLN.Size = New System.Drawing.Size(300, 21)
        Me.tbxLN.TabIndex = 15
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(61, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "Last Name:"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.PanelControl2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 26)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(437, 189)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Account"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.ForeColor = System.Drawing.Color.Red
        Me.Label9.Location = New System.Drawing.Point(91, 59)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(13, 13)
        Me.Label9.TabIndex = 7
        Me.Label9.Text = "*"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Red
        Me.Label8.Location = New System.Drawing.Point(91, 31)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(13, 13)
        Me.Label8.TabIndex = 6
        Me.Label8.Text = "*"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(338, 133)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(257, 133)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 25)
        Me.btnAdd.TabIndex = 4
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(6, 62)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 13)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Password:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(6, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(59, 13)
        Me.Label6.TabIndex = 2
        Me.Label6.Text = "Username:"
        '
        'tbxPass
        '
        Me.tbxPass.Location = New System.Drawing.Point(113, 59)
        Me.tbxPass.Name = "tbxPass"
        Me.tbxPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.tbxPass.Size = New System.Drawing.Size(300, 21)
        Me.tbxPass.TabIndex = 1
        '
        'tbxUserName
        '
        Me.tbxUserName.Location = New System.Drawing.Point(113, 31)
        Me.tbxUserName.Name = "tbxUserName"
        Me.tbxUserName.Size = New System.Drawing.Size(300, 21)
        Me.tbxUserName.TabIndex = 0
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.cbxArea)
        Me.PanelControl1.Controls.Add(Me.lblErrArea)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Controls.Add(Me.lblErrRole)
        Me.PanelControl1.Controls.Add(Me.tbxLN)
        Me.PanelControl1.Controls.Add(Me.lblErrLN)
        Me.PanelControl1.Controls.Add(Me.tbxMN)
        Me.PanelControl1.Controls.Add(Me.lblErrFN)
        Me.PanelControl1.Controls.Add(Me.cbxRole)
        Me.PanelControl1.Controls.Add(Me.Label5)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.tbxFN)
        Me.PanelControl1.Controls.Add(Me.Label4)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Location = New System.Drawing.Point(8, 6)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(432, 160)
        Me.PanelControl1.TabIndex = 24
        '
        'PanelControl2
        '
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.btnCancel)
        Me.PanelControl2.Controls.Add(Me.Label9)
        Me.PanelControl2.Controls.Add(Me.tbxUserName)
        Me.PanelControl2.Controls.Add(Me.Label8)
        Me.PanelControl2.Controls.Add(Me.tbxPass)
        Me.PanelControl2.Controls.Add(Me.Label6)
        Me.PanelControl2.Controls.Add(Me.btnAdd)
        Me.PanelControl2.Controls.Add(Me.Label7)
        Me.PanelControl2.Location = New System.Drawing.Point(6, 6)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(421, 161)
        Me.PanelControl2.TabIndex = 8
        '
        'usersAdd
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(445, 219)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "usersAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        Me.PanelControl2.PerformLayout()
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
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents tbxPass As TextBox
    Friend WithEvents tbxUserName As TextBox
    Friend WithEvents Label9 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
End Class
