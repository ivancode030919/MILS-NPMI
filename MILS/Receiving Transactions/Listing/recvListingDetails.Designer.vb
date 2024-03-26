<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class recvListingDetails
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.cbxOwnership = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbxRefDocType = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbxRefDocNum = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxEntry = New System.Windows.Forms.TextBox()
        Me.tbxEmployeeName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbxRemarks = New System.Windows.Forms.TextBox()
        Me.tbxDocType = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbxSender = New System.Windows.Forms.ComboBox()
        Me.dtpRefDate = New System.Windows.Forms.DateTimePicker()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.rb1 = New System.Windows.Forms.RadioButton()
        Me.rb2 = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.dg1 = New System.Windows.Forms.DataGridView()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.PanelControl3 = New DevExpress.XtraEditors.PanelControl()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Controls.Add(Me.cbxOwnership)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.tbxRefDocType)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.tbxRefDocNum)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tbxEntry)
        Me.GroupBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(485, 9)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(498, 179)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Reference"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(105, 147)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(68, 18)
        Me.Label7.TabIndex = 34
        Me.Label7.Text = "Courier :"
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(179, 144)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(300, 23)
        Me.TextBox1.TabIndex = 33
        '
        'cbxOwnership
        '
        Me.cbxOwnership.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxOwnership.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxOwnership.FormattingEnabled = True
        Me.cbxOwnership.Items.AddRange(New Object() {"Owned", "Consigned"})
        Me.cbxOwnership.Location = New System.Drawing.Point(179, 112)
        Me.cbxOwnership.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxOwnership.Name = "cbxOwnership"
        Me.cbxOwnership.Size = New System.Drawing.Size(300, 25)
        Me.cbxOwnership.TabIndex = 32
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(83, 115)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 18)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Ownership :"
        '
        'tbxRefDocType
        '
        Me.tbxRefDocType.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxRefDocType.Location = New System.Drawing.Point(179, 52)
        Me.tbxRefDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocType.Name = "tbxRefDocType"
        Me.tbxRefDocType.Size = New System.Drawing.Size(300, 23)
        Me.tbxRefDocType.TabIndex = 7
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(25, 85)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(148, 18)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Ref. Document No. :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(17, 55)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(156, 18)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Ref. Document Type :"
        '
        'tbxRefDocNum
        '
        Me.tbxRefDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxRefDocNum.Location = New System.Drawing.Point(179, 82)
        Me.tbxRefDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocNum.Name = "tbxRefDocNum"
        Me.tbxRefDocNum.Size = New System.Drawing.Size(300, 23)
        Me.tbxRefDocNum.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(60, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 18)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Entry Number :"
        '
        'tbxEntry
        '
        Me.tbxEntry.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEntry.Location = New System.Drawing.Point(179, 22)
        Me.tbxEntry.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxEntry.Name = "tbxEntry"
        Me.tbxEntry.ReadOnly = True
        Me.tbxEntry.Size = New System.Drawing.Size(300, 23)
        Me.tbxEntry.TabIndex = 0
        '
        'tbxEmployeeName
        '
        Me.tbxEmployeeName.Enabled = False
        Me.tbxEmployeeName.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEmployeeName.Location = New System.Drawing.Point(86, 9)
        Me.tbxEmployeeName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxEmployeeName.Name = "tbxEmployeeName"
        Me.tbxEmployeeName.ReadOnly = True
        Me.tbxEmployeeName.Size = New System.Drawing.Size(300, 23)
        Me.tbxEmployeeName.TabIndex = 13
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(5, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(75, 18)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Encoder :"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(72, 146)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(78, 18)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Remarks :"
        '
        'tbxRemarks
        '
        Me.tbxRemarks.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxRemarks.Location = New System.Drawing.Point(156, 144)
        Me.tbxRemarks.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRemarks.Multiline = True
        Me.tbxRemarks.Name = "tbxRemarks"
        Me.tbxRemarks.Size = New System.Drawing.Size(300, 24)
        Me.tbxRemarks.TabIndex = 33
        '
        'tbxDocType
        '
        Me.tbxDocType.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocType.Location = New System.Drawing.Point(156, 53)
        Me.tbxDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocType.Name = "tbxDocType"
        Me.tbxDocType.Size = New System.Drawing.Size(300, 23)
        Me.tbxDocType.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(40, 26)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(110, 18)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Receive From :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(27, 114)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(123, 18)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Document Date :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(35, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(115, 18)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Document No. :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(27, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 18)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Document Type :"
        '
        'cbxSender
        '
        Me.cbxSender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSender.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSender.FormattingEnabled = True
        Me.cbxSender.Location = New System.Drawing.Point(156, 21)
        Me.cbxSender.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSender.Name = "cbxSender"
        Me.cbxSender.Size = New System.Drawing.Size(300, 25)
        Me.cbxSender.TabIndex = 1
        '
        'dtpRefDate
        '
        Me.dtpRefDate.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpRefDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRefDate.Location = New System.Drawing.Point(156, 113)
        Me.dtpRefDate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dtpRefDate.Name = "dtpRefDate"
        Me.dtpRefDate.Size = New System.Drawing.Size(300, 23)
        Me.dtpRefDate.TabIndex = 5
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocNum.Location = New System.Drawing.Point(156, 83)
        Me.tbxDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(300, 23)
        Me.tbxDocNum.TabIndex = 4
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.tbxRemarks)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.tbxDocNum)
        Me.GroupBox4.Controls.Add(Me.dtpRefDate)
        Me.GroupBox4.Controls.Add(Me.cbxSender)
        Me.GroupBox4.Controls.Add(Me.tbxDocType)
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(8, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(471, 180)
        Me.GroupBox4.TabIndex = 6
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Details"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(29, 55)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(159, 29)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Update"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'rb1
        '
        Me.rb1.AutoSize = True
        Me.rb1.Location = New System.Drawing.Point(29, 27)
        Me.rb1.Name = "rb1"
        Me.rb1.Size = New System.Drawing.Size(73, 22)
        Me.rb1.TabIndex = 8
        Me.rb1.Text = "Enable"
        Me.rb1.UseVisualStyleBackColor = True
        '
        'rb2
        '
        Me.rb2.AutoSize = True
        Me.rb2.Checked = True
        Me.rb2.Location = New System.Drawing.Point(108, 27)
        Me.rb2.Name = "rb2"
        Me.rb2.Size = New System.Drawing.Size(77, 22)
        Me.rb2.TabIndex = 9
        Me.rb2.TabStop = True
        Me.rb2.Text = "Disable"
        Me.rb2.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.rb2)
        Me.GroupBox3.Controls.Add(Me.rb1)
        Me.GroupBox3.Controls.Add(Me.Button1)
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox3.Location = New System.Drawing.Point(989, 9)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(207, 103)
        Me.GroupBox3.TabIndex = 7
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Action"
        '
        'dg1
        '
        Me.dg1.AllowUserToOrderColumns = True
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!)
        Me.dg1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dg1.BackgroundColor = System.Drawing.Color.White
        Me.dg1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dg1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dg1.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.dg1.GridColor = System.Drawing.Color.Silver
        Me.dg1.Location = New System.Drawing.Point(0, 0)
        Me.dg1.Name = "dg1"
        Me.dg1.RowHeadersVisible = False
        Me.dg1.Size = New System.Drawing.Size(1191, 338)
        Me.dg1.TabIndex = 8
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.dg1)
        Me.PanelControl1.Location = New System.Drawing.Point(5, 195)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(1191, 338)
        Me.PanelControl1.TabIndex = 9
        '
        'PanelControl2
        '
        Me.PanelControl2.Controls.Add(Me.Button2)
        Me.PanelControl2.Location = New System.Drawing.Point(1027, 539)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(169, 40)
        Me.PanelControl2.TabIndex = 10
        '
        'Button2
        '
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button2.Location = New System.Drawing.Point(2, 2)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(165, 36)
        Me.Button2.TabIndex = 10
        Me.Button2.Text = "Print"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'PanelControl3
        '
        Me.PanelControl3.Controls.Add(Me.tbxEmployeeName)
        Me.PanelControl3.Controls.Add(Me.Label5)
        Me.PanelControl3.Location = New System.Drawing.Point(5, 539)
        Me.PanelControl3.Name = "PanelControl3"
        Me.PanelControl3.Size = New System.Drawing.Size(395, 40)
        Me.PanelControl3.TabIndex = 11
        '
        'recvListingDetails
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.PanelControl3)
        Me.Controls.Add(Me.PanelControl2)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "recvListingDetails"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " "
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.dg1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        CType(Me.PanelControl3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl3.ResumeLayout(False)
        Me.PanelControl3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label9 As Label
    Friend WithEvents tbxRemarks As TextBox
    Friend WithEvents cbxOwnership As ComboBox
    Friend WithEvents Label8 As Label
    Friend WithEvents tbxRefDocType As TextBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label11 As Label
    Friend WithEvents tbxRefDocNum As TextBox
    Friend WithEvents tbxDocType As TextBox
    Friend WithEvents Label6 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cbxSender As ComboBox
    Friend WithEvents dtpRefDate As DateTimePicker
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents tbxEntry As TextBox
    Friend WithEvents tbxEmployeeName As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Button1 As Button
    Friend WithEvents rb1 As RadioButton
    Friend WithEvents rb2 As RadioButton
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents dg1 As DataGridView
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents Button2 As Button
    Friend WithEvents PanelControl3 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents Label7 As Label
    Friend WithEvents TextBox1 As TextBox
End Class
