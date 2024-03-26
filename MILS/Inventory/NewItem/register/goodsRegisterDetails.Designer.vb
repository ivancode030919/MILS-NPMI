<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class goodsRegisterDetails
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
        Me.tbxEntry = New System.Windows.Forms.TextBox()
        Me.dtpDateEnc = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.tbxRemarks = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lvDetails = New System.Windows.Forms.ListView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxUser = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbxApproved = New System.Windows.Forms.TextBox()
        Me.bntClose = New System.Windows.Forms.Button()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbxEntry
        '
        Me.tbxEntry.Location = New System.Drawing.Point(552, 12)
        Me.tbxEntry.Name = "tbxEntry"
        Me.tbxEntry.Size = New System.Drawing.Size(220, 22)
        Me.tbxEntry.TabIndex = 3
        '
        'dtpDateEnc
        '
        Me.dtpDateEnc.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDateEnc.Location = New System.Drawing.Point(97, 12)
        Me.dtpDateEnc.Name = "dtpDateEnc"
        Me.dtpDateEnc.Size = New System.Drawing.Size(250, 22)
        Me.dtpDateEnc.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(38, 17)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(46, 17)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Date :"
        '
        'tbxRemarks
        '
        Me.tbxRemarks.Location = New System.Drawing.Point(97, 40)
        Me.tbxRemarks.Name = "tbxRemarks"
        Me.tbxRemarks.Size = New System.Drawing.Size(250, 22)
        Me.tbxRemarks.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(11, 43)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Remarks :"
        '
        'lvDetails
        '
        Me.lvDetails.HideSelection = False
        Me.lvDetails.Location = New System.Drawing.Point(12, 68)
        Me.lvDetails.Name = "lvDetails"
        Me.lvDetails.Size = New System.Drawing.Size(760, 267)
        Me.lvDetails.TabIndex = 4
        Me.lvDetails.UseCompatibleStateImageBehavior = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(62, 345)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(45, 17)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "User :"
        '
        'tbxUser
        '
        Me.tbxUser.Location = New System.Drawing.Point(109, 341)
        Me.tbxUser.Name = "tbxUser"
        Me.tbxUser.Size = New System.Drawing.Size(250, 21)
        Me.tbxUser.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(435, 17)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(103, 17)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Entry Number :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(12, 372)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(95, 17)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "Approved By :"
        '
        'tbxApproved
        '
        Me.tbxApproved.Location = New System.Drawing.Point(109, 369)
        Me.tbxApproved.Name = "tbxApproved"
        Me.tbxApproved.Size = New System.Drawing.Size(250, 21)
        Me.tbxApproved.TabIndex = 6
        '
        'bntClose
        '
        Me.bntClose.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Bold)
        Me.bntClose.Location = New System.Drawing.Point(622, 363)
        Me.bntClose.Name = "bntClose"
        Me.bntClose.Size = New System.Drawing.Size(150, 30)
        Me.bntClose.TabIndex = 7
        Me.bntClose.Text = "Close"
        Me.bntClose.UseVisualStyleBackColor = True
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.tbxApproved)
        Me.PanelControl1.Controls.Add(Me.tbxUser)
        Me.PanelControl1.Controls.Add(Me.Label5)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(784, 401)
        Me.PanelControl1.TabIndex = 22
        '
        'goodsRegisterDetails
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 401)
        Me.Controls.Add(Me.bntClose)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.lvDetails)
        Me.Controls.Add(Me.tbxEntry)
        Me.Controls.Add(Me.dtpDateEnc)
        Me.Controls.Add(Me.tbxRemarks)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.PanelControl1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "goodsRegisterDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Details"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tbxEntry As TextBox
    Friend WithEvents dtpDateEnc As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents tbxRemarks As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents lvDetails As ListView
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxUser As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents tbxApproved As TextBox
    Friend WithEvents bntClose As Button
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
