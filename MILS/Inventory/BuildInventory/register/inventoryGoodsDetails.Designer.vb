<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class inventoryGoodsDetails
    Inherits DevExpress.XtraEditors.XtraForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.tbxApprvdBy = New System.Windows.Forms.TextBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbxRemarks = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxEntryNumber = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbxOwnership = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.tbxUser = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.lvDetails = New System.Windows.Forms.ListView()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbxApprvdBy
        '
        Me.tbxApprvdBy.Enabled = False
        Me.tbxApprvdBy.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxApprvdBy.Location = New System.Drawing.Point(870, 76)
        Me.tbxApprvdBy.Name = "tbxApprvdBy"
        Me.tbxApprvdBy.Size = New System.Drawing.Size(300, 23)
        Me.tbxApprvdBy.TabIndex = 40
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(172, 110)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(17, 18)
        Me.lblErr.TabIndex = 1
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(750, 81)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 18)
        Me.Label6.TabIndex = 41
        Me.Label6.Text = "Approved By :"
        '
        'tbxRemarks
        '
        Me.tbxRemarks.Enabled = False
        Me.tbxRemarks.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxRemarks.Location = New System.Drawing.Point(870, 49)
        Me.tbxRemarks.Name = "tbxRemarks"
        Me.tbxRemarks.Size = New System.Drawing.Size(300, 23)
        Me.tbxRemarks.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(774, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Remarks :"
        '
        'tbxEntryNumber
        '
        Me.tbxEntryNumber.Enabled = False
        Me.tbxEntryNumber.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEntryNumber.Location = New System.Drawing.Point(870, 22)
        Me.tbxEntryNumber.Name = "tbxEntryNumber"
        Me.tbxEntryNumber.Size = New System.Drawing.Size(300, 23)
        Me.tbxEntryNumber.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(739, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Entry Number :"
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Enabled = False
        Me.tbxDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocNum.Location = New System.Drawing.Point(175, 22)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(300, 23)
        Me.tbxDocNum.TabIndex = 36
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(146, 18)
        Me.Label3.TabIndex = 37
        Me.Label3.Text = "Document Number :"
        '
        'cbxOwnership
        '
        Me.cbxOwnership.Enabled = False
        Me.cbxOwnership.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxOwnership.FormattingEnabled = True
        Me.cbxOwnership.Items.AddRange(New Object() {"Owned", "Consigned"})
        Me.cbxOwnership.Location = New System.Drawing.Point(175, 76)
        Me.cbxOwnership.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxOwnership.Name = "cbxOwnership"
        Me.cbxOwnership.Size = New System.Drawing.Size(300, 25)
        Me.cbxOwnership.TabIndex = 36
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(79, 83)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 18)
        Me.Label8.TabIndex = 35
        Me.Label8.Text = "Ownership :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(46, 54)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(123, 18)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Document Date :"
        '
        'dtpDate
        '
        Me.dtpDate.Enabled = False
        Me.dtpDate.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpDate.Location = New System.Drawing.Point(175, 49)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Size = New System.Drawing.Size(300, 23)
        Me.dtpDate.TabIndex = 2
        '
        'tbxUser
        '
        Me.tbxUser.Enabled = False
        Me.tbxUser.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxUser.Location = New System.Drawing.Point(870, 105)
        Me.tbxUser.Name = "tbxUser"
        Me.tbxUser.Size = New System.Drawing.Size(300, 23)
        Me.tbxUser.TabIndex = 38
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(777, 110)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 18)
        Me.Label5.TabIndex = 39
        Me.Label5.Text = "Encoder :"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.lvDetails)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(5, 148)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1191, 429)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Contents"
        '
        'lvDetails
        '
        Me.lvDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvDetails.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvDetails.HideSelection = False
        Me.lvDetails.Location = New System.Drawing.Point(3, 19)
        Me.lvDetails.Name = "lvDetails"
        Me.lvDetails.Size = New System.Drawing.Size(1185, 407)
        Me.lvDetails.TabIndex = 0
        Me.lvDetails.UseCompatibleStateImageBehavior = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblErr)
        Me.GroupBox3.Controls.Add(Me.tbxApprvdBy)
        Me.GroupBox3.Controls.Add(Me.tbxDocNum)
        Me.GroupBox3.Controls.Add(Me.Label6)
        Me.GroupBox3.Controls.Add(Me.cbxOwnership)
        Me.GroupBox3.Controls.Add(Me.tbxRemarks)
        Me.GroupBox3.Controls.Add(Me.tbxUser)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.tbxEntryNumber)
        Me.GroupBox3.Controls.Add(Me.dtpDate)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label8)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1191, 137)
        Me.GroupBox3.TabIndex = 40
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Details"
        '
        'inventoryGoodsDetails
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox4)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "inventoryGoodsDetails"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Details"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label4 As Label
    Friend WithEvents dtpDate As DateTimePicker
    Friend WithEvents tbxRemarks As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxEntryNumber As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblErr As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents lvDetails As ListView
    Friend WithEvents cbxOwnership As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tbxApprvdBy As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents tbxUser As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents GroupBox3 As GroupBox
End Class
