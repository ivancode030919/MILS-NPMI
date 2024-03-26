<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class recvListing
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
        Me.lvRecvListing = New System.Windows.Forms.ListView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbxSender = New System.Windows.Forms.ComboBox()
        Me.rdIndi = New System.Windows.Forms.RadioButton()
        Me.rdAll = New System.Windows.Forms.RadioButton()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpfrom = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblError = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.tbxDocType = New System.Windows.Forms.TextBox()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.gbArea = New System.Windows.Forms.GroupBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.rdSelectCategArea = New System.Windows.Forms.RadioButton()
        Me.rdAllArea = New System.Windows.Forms.RadioButton()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.gbArea.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvRecvListing
        '
        Me.lvRecvListing.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvRecvListing.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!)
        Me.lvRecvListing.HideSelection = False
        Me.lvRecvListing.Location = New System.Drawing.Point(0, 0)
        Me.lvRecvListing.Name = "lvRecvListing"
        Me.lvRecvListing.Size = New System.Drawing.Size(1191, 481)
        Me.lvRecvListing.TabIndex = 5
        Me.lvRecvListing.UseCompatibleStateImageBehavior = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbxSender)
        Me.GroupBox1.Controls.Add(Me.rdIndi)
        Me.GroupBox1.Controls.Add(Me.rdAll)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(621, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(260, 82)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Source"
        '
        'cbxSender
        '
        Me.cbxSender.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxSender.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxSender.FormattingEnabled = True
        Me.cbxSender.Location = New System.Drawing.Point(6, 45)
        Me.cbxSender.Name = "cbxSender"
        Me.cbxSender.Size = New System.Drawing.Size(248, 24)
        Me.cbxSender.TabIndex = 3
        '
        'rdIndi
        '
        Me.rdIndi.AutoSize = True
        Me.rdIndi.Location = New System.Drawing.Point(101, 22)
        Me.rdIndi.Name = "rdIndi"
        Me.rdIndi.Size = New System.Drawing.Size(127, 21)
        Me.rdIndi.TabIndex = 2
        Me.rdIndi.Text = "Select Source"
        Me.rdIndi.UseVisualStyleBackColor = True
        '
        'rdAll
        '
        Me.rdAll.AutoSize = True
        Me.rdAll.Checked = True
        Me.rdAll.Location = New System.Drawing.Point(36, 22)
        Me.rdAll.Name = "rdAll"
        Me.rdAll.Size = New System.Drawing.Size(44, 21)
        Me.rdAll.TabIndex = 1
        Me.rdAll.TabStop = True
        Me.rdAll.Text = "All"
        Me.rdAll.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(23, 57)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "To :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 24)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 17)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "From :"
        '
        'dtpTo
        '
        Me.dtpTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(68, 51)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(226, 23)
        Me.dtpTo.TabIndex = 9
        Me.dtpTo.Value = New Date(2023, 5, 26, 9, 33, 24, 0)
        '
        'dtpfrom
        '
        Me.dtpfrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpfrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpfrom.Location = New System.Drawing.Point(68, 22)
        Me.dtpfrom.Name = "dtpfrom"
        Me.dtpfrom.Size = New System.Drawing.Size(226, 23)
        Me.dtpfrom.TabIndex = 8
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.dtpfrom)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(5, 8)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(311, 82)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Tag = "4"
        Me.GroupBox2.Text = "Date Range"
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(599, 330)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(16, 20)
        Me.lblError.TabIndex = 8
        Me.lblError.Text = "*"
        Me.lblError.Visible = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.tbxDocType)
        Me.GroupBox4.Controls.Add(Me.tbxDocNum)
        Me.GroupBox4.Controls.Add(Me.lblNumber)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(322, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(293, 82)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Document"
        '
        'tbxDocType
        '
        Me.tbxDocType.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocType.Location = New System.Drawing.Point(86, 18)
        Me.tbxDocType.Name = "tbxDocType"
        Me.tbxDocType.Size = New System.Drawing.Size(191, 23)
        Me.tbxDocType.TabIndex = 5
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocNum.Location = New System.Drawing.Point(86, 45)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(191, 23)
        Me.tbxDocNum.TabIndex = 6
        '
        'lblNumber
        '
        Me.lblNumber.AutoSize = True
        Me.lblNumber.Location = New System.Drawing.Point(6, 51)
        Me.lblNumber.Name = "lblNumber"
        Me.lblNumber.Size = New System.Drawing.Size(74, 17)
        Me.lblNumber.TabIndex = 20
        Me.lblNumber.Text = "Number :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(26, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 17)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Type :"
        '
        'gbArea
        '
        Me.gbArea.Controls.Add(Me.cbxArea)
        Me.gbArea.Controls.Add(Me.rdSelectCategArea)
        Me.gbArea.Controls.Add(Me.rdAllArea)
        Me.gbArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbArea.Location = New System.Drawing.Point(887, 8)
        Me.gbArea.Name = "gbArea"
        Me.gbArea.Size = New System.Drawing.Size(218, 82)
        Me.gbArea.TabIndex = 4
        Me.gbArea.TabStop = False
        Me.gbArea.Tag = "4"
        Me.gbArea.Text = "Area"
        Me.gbArea.Visible = False
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(6, 44)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(206, 24)
        Me.cbxArea.TabIndex = 6
        '
        'rdSelectCategArea
        '
        Me.rdSelectCategArea.AutoSize = True
        Me.rdSelectCategArea.Location = New System.Drawing.Point(102, 22)
        Me.rdSelectCategArea.Name = "rdSelectCategArea"
        Me.rdSelectCategArea.Size = New System.Drawing.Size(110, 21)
        Me.rdSelectCategArea.TabIndex = 5
        Me.rdSelectCategArea.Text = "Select Area"
        Me.rdSelectCategArea.UseVisualStyleBackColor = True
        '
        'rdAllArea
        '
        Me.rdAllArea.AutoSize = True
        Me.rdAllArea.Checked = True
        Me.rdAllArea.Location = New System.Drawing.Point(38, 22)
        Me.rdAllArea.Name = "rdAllArea"
        Me.rdAllArea.Size = New System.Drawing.Size(44, 21)
        Me.rdAllArea.TabIndex = 4
        Me.rdAllArea.TabStop = True
        Me.rdAllArea.Text = "All"
        Me.rdAllArea.UseVisualStyleBackColor = True
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.lvRecvListing)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelControl1.Location = New System.Drawing.Point(5, 96)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(1191, 481)
        Me.PanelControl1.TabIndex = 9
        '
        'recvListing
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.gbArea)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "recvListing"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Receive Register"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.gbArea.ResumeLayout(False)
        Me.gbArea.PerformLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lvRecvListing As ListView
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents rdIndi As RadioButton
    Friend WithEvents rdAll As RadioButton
    Friend WithEvents cbxSender As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents dtpfrom As DateTimePicker
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents lblError As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents tbxDocType As TextBox
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents lblNumber As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents gbArea As GroupBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents rdSelectCategArea As RadioButton
    Friend WithEvents rdAllArea As RadioButton
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
