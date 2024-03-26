<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class reserveRelease
    Inherits System.Windows.Forms.Form

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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.tbxRemarks = New System.Windows.Forms.TextBox()
        Me.cbxOwnership = New System.Windows.Forms.ComboBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.tbxRefDocType = New System.Windows.Forms.TextBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.tbxRefDocNum = New System.Windows.Forms.TextBox()
        Me.tbxDocType = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbxReceiver = New System.Windows.Forms.ComboBox()
        Me.dtpRefDate = New System.Windows.Forms.DateTimePicker()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.tbxEntry = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.dgvRels = New System.Windows.Forms.DataGridView()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.dgvRels, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label16)
        Me.GroupBox1.Controls.Add(Me.Label14)
        Me.GroupBox1.Controls.Add(Me.Label13)
        Me.GroupBox1.Controls.Add(Me.Label12)
        Me.GroupBox1.Controls.Add(Me.Label9)
        Me.GroupBox1.Controls.Add(Me.tbxRemarks)
        Me.GroupBox1.Controls.Add(Me.cbxOwnership)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.tbxRefDocType)
        Me.GroupBox1.Controls.Add(Me.Label10)
        Me.GroupBox1.Controls.Add(Me.Label11)
        Me.GroupBox1.Controls.Add(Me.tbxRefDocNum)
        Me.GroupBox1.Controls.Add(Me.tbxDocType)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cbxReceiver)
        Me.GroupBox1.Controls.Add(Me.dtpRefDate)
        Me.GroupBox1.Controls.Add(Me.tbxDocNum)
        Me.GroupBox1.Controls.Add(Me.tbxEntry)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox1.Size = New System.Drawing.Size(1284, 173)
        Me.GroupBox1.TabIndex = 4
        Me.GroupBox1.TabStop = False
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.ForeColor = System.Drawing.Color.Red
        Me.Label16.Location = New System.Drawing.Point(950, 43)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(16, 17)
        Me.Label16.TabIndex = 39
        Me.Label16.Text = "*"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.ForeColor = System.Drawing.Color.Red
        Me.Label14.Location = New System.Drawing.Point(136, 75)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(16, 17)
        Me.Label14.TabIndex = 37
        Me.Label14.Text = "*"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.ForeColor = System.Drawing.Color.Red
        Me.Label13.Location = New System.Drawing.Point(136, 45)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(16, 17)
        Me.Label13.TabIndex = 36
        Me.Label13.Text = "*"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.ForeColor = System.Drawing.Color.Red
        Me.Label12.Location = New System.Drawing.Point(136, 13)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(16, 17)
        Me.Label12.TabIndex = 35
        Me.Label12.Text = "*"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(14, 139)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(67, 17)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Remarks : "
        '
        'tbxRemarks
        '
        Me.tbxRemarks.Location = New System.Drawing.Point(158, 136)
        Me.tbxRemarks.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRemarks.Name = "tbxRemarks"
        Me.tbxRemarks.Size = New System.Drawing.Size(1114, 22)
        Me.tbxRemarks.TabIndex = 33
        '
        'cbxOwnership
        '
        Me.cbxOwnership.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxOwnership.FormattingEnabled = True
        Me.cbxOwnership.Items.AddRange(New Object() {"Owned", "Consigned"})
        Me.cbxOwnership.Location = New System.Drawing.Point(972, 103)
        Me.cbxOwnership.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxOwnership.Name = "cbxOwnership"
        Me.cbxOwnership.Size = New System.Drawing.Size(300, 25)
        Me.cbxOwnership.TabIndex = 32
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(808, 106)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(69, 17)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Ownership:"
        '
        'tbxRefDocType
        '
        Me.tbxRefDocType.Location = New System.Drawing.Point(972, 43)
        Me.tbxRefDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocType.Name = "tbxRefDocType"
        Me.tbxRefDocType.Size = New System.Drawing.Size(300, 22)
        Me.tbxRefDocType.TabIndex = 7
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(808, 76)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(115, 17)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Ref. Document No.:"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(808, 46)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(121, 17)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Ref. Document Type:"
        '
        'tbxRefDocNum
        '
        Me.tbxRefDocNum.Location = New System.Drawing.Point(972, 73)
        Me.tbxRefDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocNum.Name = "tbxRefDocNum"
        Me.tbxRefDocNum.Size = New System.Drawing.Size(300, 22)
        Me.tbxRefDocNum.TabIndex = 8
        '
        'tbxDocType
        '
        Me.tbxDocType.Location = New System.Drawing.Point(158, 45)
        Me.tbxDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocType.Name = "tbxDocType"
        Me.tbxDocType.Size = New System.Drawing.Size(300, 22)
        Me.tbxDocType.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(14, 19)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 17)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Release To:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(14, 107)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(97, 17)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Document Date:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(14, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 17)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Document No. : "
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(14, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(99, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Document Type: "
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(808, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(86, 17)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Entry Number:"
        '
        'cbxReceiver
        '
        Me.cbxReceiver.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxReceiver.FormattingEnabled = True
        Me.cbxReceiver.Location = New System.Drawing.Point(158, 13)
        Me.cbxReceiver.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxReceiver.Name = "cbxReceiver"
        Me.cbxReceiver.Size = New System.Drawing.Size(300, 25)
        Me.cbxReceiver.TabIndex = 1
        '
        'dtpRefDate
        '
        Me.dtpRefDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRefDate.Location = New System.Drawing.Point(158, 105)
        Me.dtpRefDate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dtpRefDate.Name = "dtpRefDate"
        Me.dtpRefDate.Size = New System.Drawing.Size(300, 22)
        Me.dtpRefDate.TabIndex = 5
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Location = New System.Drawing.Point(158, 75)
        Me.tbxDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(300, 22)
        Me.tbxDocNum.TabIndex = 4
        '
        'tbxEntry
        '
        Me.tbxEntry.Location = New System.Drawing.Point(972, 13)
        Me.tbxEntry.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxEntry.Name = "tbxEntry"
        Me.tbxEntry.ReadOnly = True
        Me.tbxEntry.Size = New System.Drawing.Size(300, 22)
        Me.tbxEntry.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label5)
        Me.GroupBox2.Controls.Add(Me.Label15)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.dgvRels)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 181)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.GroupBox2.Size = New System.Drawing.Size(1260, 409)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Contents"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(68, -4)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(16, 17)
        Me.Label5.TabIndex = 50
        Me.Label5.Text = "*"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.ForeColor = System.Drawing.Color.Red
        Me.Label15.Location = New System.Drawing.Point(1170, 28)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(16, 17)
        Me.Label15.TabIndex = 44
        Me.Label15.Text = "*"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(166, 28)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(16, 17)
        Me.Label7.TabIndex = 43
        Me.Label7.Text = "*"
        '
        'dgvRels
        '
        Me.dgvRels.ColumnHeadersHeight = 45
        Me.dgvRels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvRels.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRels.Location = New System.Drawing.Point(3, 19)
        Me.dgvRels.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvRels.Name = "dgvRels"
        Me.dgvRels.RowHeadersWidth = 50
        Me.dgvRels.Size = New System.Drawing.Size(1254, 386)
        Me.dgvRels.TabIndex = 8
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(966, 598)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(150, 30)
        Me.btnCancel.TabIndex = 49
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(12, 605)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(16, 17)
        Me.lblErr.TabIndex = 48
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(1122, 598)
        Me.btnAdd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(150, 30)
        Me.btnAdd.TabIndex = 47
        Me.btnAdd.Text = "Record"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'reserveRelease
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1284, 641)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.lblErr)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "reserveRelease"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Reserve"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.dgvRels, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Label14 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
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
    Friend WithEvents cbxReceiver As ComboBox
    Friend WithEvents dtpRefDate As DateTimePicker
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents tbxEntry As TextBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label15 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents dgvRels As DataGridView
    Friend WithEvents btnCancel As Button
    Friend WithEvents lblErr As Label
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label5 As Label
End Class
