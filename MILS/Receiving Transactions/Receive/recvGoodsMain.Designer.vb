<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class recvGoodsMain
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
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
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
        Me.cbxSender = New System.Windows.Forms.ComboBox()
        Me.dtpRefDate = New System.Windows.Forms.DateTimePicker()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.tbxEntry = New System.Windows.Forms.TextBox()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.dgvRecv = New System.Windows.Forms.DataGridView()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        CType(Me.dgvRecv, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox4.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label14.ForeColor = System.Drawing.Color.Red
        Me.Label14.Location = New System.Drawing.Point(472, 89)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(17, 18)
        Me.Label14.TabIndex = 37
        Me.Label14.Text = "*"
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label13.ForeColor = System.Drawing.Color.Red
        Me.Label13.Location = New System.Drawing.Point(472, 59)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(17, 18)
        Me.Label13.TabIndex = 36
        Me.Label13.Text = "*"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label12.ForeColor = System.Drawing.Color.Red
        Me.Label12.Location = New System.Drawing.Point(472, 27)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(17, 18)
        Me.Label12.TabIndex = 35
        Me.Label12.Text = "*"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label9.Location = New System.Drawing.Point(75, 145)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(78, 18)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Remarks :"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbxOwnership
        '
        Me.cbxOwnership.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxOwnership.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxOwnership.FormattingEnabled = True
        Me.cbxOwnership.Items.AddRange(New Object() {"Owned", "Consigned"})
        Me.cbxOwnership.Location = New System.Drawing.Point(862, 114)
        Me.cbxOwnership.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxOwnership.Name = "cbxOwnership"
        Me.cbxOwnership.Size = New System.Drawing.Size(300, 25)
        Me.cbxOwnership.TabIndex = 32
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label8.Location = New System.Drawing.Point(766, 120)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(90, 18)
        Me.Label8.TabIndex = 31
        Me.Label8.Text = "Ownership :"
        '
        'tbxRefDocType
        '
        Me.tbxRefDocType.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxRefDocType.Location = New System.Drawing.Point(862, 50)
        Me.tbxRefDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocType.Name = "tbxRefDocType"
        Me.tbxRefDocType.Size = New System.Drawing.Size(300, 23)
        Me.tbxRefDocType.TabIndex = 7
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label10.Location = New System.Drawing.Point(677, 88)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(179, 18)
        Me.Label10.TabIndex = 28
        Me.Label10.Text = "Ref. Document Number :"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label11.Location = New System.Drawing.Point(700, 55)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(156, 18)
        Me.Label11.TabIndex = 27
        Me.Label11.Text = "Ref. Document Type :"
        '
        'tbxRefDocNum
        '
        Me.tbxRefDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxRefDocNum.Location = New System.Drawing.Point(862, 84)
        Me.tbxRefDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxRefDocNum.Name = "tbxRefDocNum"
        Me.tbxRefDocNum.Size = New System.Drawing.Size(300, 23)
        Me.tbxRefDocNum.TabIndex = 8
        '
        'tbxDocType
        '
        Me.tbxDocType.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxDocType.Location = New System.Drawing.Point(159, 49)
        Me.tbxDocType.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocType.Name = "tbxDocType"
        Me.tbxDocType.Size = New System.Drawing.Size(300, 23)
        Me.tbxDocType.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(43, 21)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(110, 18)
        Me.Label6.TabIndex = 16
        Me.Label6.Text = "Receive From :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(30, 117)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(123, 18)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Document Date :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(7, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(146, 18)
        Me.Label3.TabIndex = 13
        Me.Label3.Text = "Document Number :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(30, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 18)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Document Type :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(743, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 18)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Entry Number :"
        '
        'cbxSender
        '
        Me.cbxSender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSender.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxSender.FormattingEnabled = True
        Me.cbxSender.Location = New System.Drawing.Point(159, 18)
        Me.cbxSender.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.cbxSender.Name = "cbxSender"
        Me.cbxSender.Size = New System.Drawing.Size(300, 25)
        Me.cbxSender.TabIndex = 1
        '
        'dtpRefDate
        '
        Me.dtpRefDate.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.dtpRefDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpRefDate.Location = New System.Drawing.Point(159, 112)
        Me.dtpRefDate.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dtpRefDate.Name = "dtpRefDate"
        Me.dtpRefDate.Size = New System.Drawing.Size(300, 23)
        Me.dtpRefDate.TabIndex = 5
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxDocNum.Location = New System.Drawing.Point(159, 82)
        Me.tbxDocNum.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(300, 23)
        Me.tbxDocNum.TabIndex = 4
        '
        'tbxEntry
        '
        Me.tbxEntry.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxEntry.Location = New System.Drawing.Point(862, 19)
        Me.tbxEntry.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxEntry.Name = "tbxEntry"
        Me.tbxEntry.ReadOnly = True
        Me.tbxEntry.Size = New System.Drawing.Size(300, 23)
        Me.tbxEntry.TabIndex = 0
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.BackColor = System.Drawing.Color.White
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Red
        Me.Label16.Location = New System.Drawing.Point(691, 12)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(14, 17)
        Me.Label16.TabIndex = 47
        Me.Label16.Text = "*"
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.BackColor = System.Drawing.Color.White
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Red
        Me.Label20.Location = New System.Drawing.Point(1167, 12)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(14, 17)
        Me.Label20.TabIndex = 46
        Me.Label20.Text = "*"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Red
        Me.Label7.Location = New System.Drawing.Point(181, 12)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(14, 17)
        Me.Label7.TabIndex = 43
        Me.Label7.Text = "*"
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.BackColor = System.Drawing.Color.White
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.ForeColor = System.Drawing.Color.Red
        Me.Label19.Location = New System.Drawing.Point(892, 12)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(14, 17)
        Me.Label19.TabIndex = 45
        Me.Label19.Text = "*"
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.BackColor = System.Drawing.Color.White
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.Red
        Me.Label15.Location = New System.Drawing.Point(796, 12)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(14, 17)
        Me.Label15.TabIndex = 44
        Me.Label15.Text = "*"
        '
        'dgvRecv
        '
        Me.dgvRecv.AllowUserToResizeColumns = False
        Me.dgvRecv.BackgroundColor = System.Drawing.Color.WhiteSmoke
        Me.dgvRecv.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.dgvRecv.ColumnHeadersHeight = 45
        Me.dgvRecv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvRecv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRecv.GridColor = System.Drawing.Color.Silver
        Me.dgvRecv.Location = New System.Drawing.Point(0, 0)
        Me.dgvRecv.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.dgvRecv.Name = "dgvRecv"
        Me.dgvRecv.RowHeadersWidth = 50
        Me.dgvRecv.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvRecv.Size = New System.Drawing.Size(1187, 349)
        Me.dgvRecv.TabIndex = 9
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.ForeColor = System.Drawing.Color.Red
        Me.Label5.Location = New System.Drawing.Point(1168, 91)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(17, 18)
        Me.Label5.TabIndex = 42
        Me.Label5.Text = "*"
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(1168, 58)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(17, 18)
        Me.lblErr.TabIndex = 25
        Me.lblErr.Text = "*"
        '
        'btnAdd
        '
        Me.btnAdd.Font = New System.Drawing.Font("Franklin Gothic Book", 13.0!, System.Drawing.FontStyle.Bold)
        Me.btnAdd.Location = New System.Drawing.Point(952, 542)
        Me.btnAdd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(244, 30)
        Me.btnAdd.TabIndex = 12
        Me.btnAdd.Text = "Record"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label18)
        Me.GroupBox4.Controls.Add(Me.TextBox1)
        Me.GroupBox4.Controls.Add(Me.Label17)
        Me.GroupBox4.Controls.Add(Me.ComboBox1)
        Me.GroupBox4.Controls.Add(Me.lblErr)
        Me.GroupBox4.Controls.Add(Me.tbxDocNum)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.cbxOwnership)
        Me.GroupBox4.Controls.Add(Me.dtpRefDate)
        Me.GroupBox4.Controls.Add(Me.cbxSender)
        Me.GroupBox4.Controls.Add(Me.tbxRefDocType)
        Me.GroupBox4.Controls.Add(Me.Label14)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.tbxRefDocNum)
        Me.GroupBox4.Controls.Add(Me.Label13)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Controls.Add(Me.tbxEntry)
        Me.GroupBox4.Controls.Add(Me.Label8)
        Me.GroupBox4.Controls.Add(Me.Label12)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.Label10)
        Me.GroupBox4.Controls.Add(Me.Label9)
        Me.GroupBox4.Controls.Add(Me.Label11)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Controls.Add(Me.tbxDocType)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(1191, 172)
        Me.GroupBox4.TabIndex = 47
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Details"
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label18.Location = New System.Drawing.Point(788, 146)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(68, 18)
        Me.Label18.TabIndex = 46
        Me.Label18.Text = "Courier :"
        Me.Label18.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.TextBox1.Location = New System.Drawing.Point(862, 143)
        Me.TextBox1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(300, 23)
        Me.TextBox1.TabIndex = 45
        Me.TextBox1.Visible = False
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label17.ForeColor = System.Drawing.Color.Red
        Me.Label17.Location = New System.Drawing.Point(472, 145)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(17, 18)
        Me.Label17.TabIndex = 44
        Me.Label17.Text = "*"
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(159, 143)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(300, 25)
        Me.ComboBox1.TabIndex = 43
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.Label16)
        Me.Panel1.Controls.Add(Me.Label20)
        Me.Panel1.Controls.Add(Me.Label19)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.dgvRecv)
        Me.Panel1.Location = New System.Drawing.Point(5, 183)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1191, 353)
        Me.Panel1.TabIndex = 48
        '
        'recvGoodsMain
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.btnAdd)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "recvGoodsMain"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Receive Goods"
        CType(Me.dgvRecv, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label9 As Label
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
    Friend WithEvents dgvRecv As DataGridView
    Friend WithEvents lblErr As Label
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label14 As Label
    Friend WithEvents Label13 As Label
    Friend WithEvents Label12 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label15 As Label
    Friend WithEvents Label19 As Label
    Friend WithEvents Label20 As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Label16 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label17 As Label
    Friend WithEvents Label18 As Label
    Friend WithEvents TextBox1 As TextBox
End Class
