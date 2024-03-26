<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class releaseRegister
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
        Me.lvRelsListing = New System.Windows.Forms.ListView()
        Me.lblError = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.tbxDocType = New System.Windows.Forms.TextBox()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.lblNumber = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpfrom = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cbxReceiver = New System.Windows.Forms.ComboBox()
        Me.rdIndi = New System.Windows.Forms.RadioButton()
        Me.rdAll = New System.Windows.Forms.RadioButton()
        Me.gbArea = New System.Windows.Forms.GroupBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.rdSelectCategArea = New System.Windows.Forms.RadioButton()
        Me.rdAllArea = New System.Windows.Forms.RadioButton()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gbArea.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvRelsListing
        '
        Me.lvRelsListing.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.lvRelsListing.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvRelsListing.HideSelection = False
        Me.lvRelsListing.Location = New System.Drawing.Point(5, 100)
        Me.lvRelsListing.Name = "lvRelsListing"
        Me.lvRelsListing.Size = New System.Drawing.Size(1191, 477)
        Me.lvRelsListing.TabIndex = 15
        Me.lvRelsListing.UseCompatibleStateImageBehavior = False
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(591, 312)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(16, 20)
        Me.lblError.TabIndex = 22
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
        Me.GroupBox4.Location = New System.Drawing.Point(317, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(290, 86)
        Me.GroupBox4.TabIndex = 2
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Document"
        '
        'tbxDocType
        '
        Me.tbxDocType.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocType.Location = New System.Drawing.Point(88, 21)
        Me.tbxDocType.Name = "tbxDocType"
        Me.tbxDocType.Size = New System.Drawing.Size(184, 23)
        Me.tbxDocType.TabIndex = 5
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocNum.Location = New System.Drawing.Point(88, 49)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(184, 23)
        Me.tbxDocNum.TabIndex = 6
        '
        'lblNumber
        '
        Me.lblNumber.AutoSize = True
        Me.lblNumber.Location = New System.Drawing.Point(8, 51)
        Me.lblNumber.Name = "lblNumber"
        Me.lblNumber.Size = New System.Drawing.Size(74, 17)
        Me.lblNumber.TabIndex = 20
        Me.lblNumber.Text = "Number :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(28, 25)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(54, 17)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Type :"
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
        Me.GroupBox2.Size = New System.Drawing.Size(306, 86)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Tag = "4"
        Me.GroupBox2.Text = "Date Range"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        'dtpfrom
        '
        Me.dtpfrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpfrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpfrom.Location = New System.Drawing.Point(68, 22)
        Me.dtpfrom.Name = "dtpfrom"
        Me.dtpfrom.Size = New System.Drawing.Size(226, 23)
        Me.dtpfrom.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbxReceiver)
        Me.GroupBox1.Controls.Add(Me.rdIndi)
        Me.GroupBox1.Controls.Add(Me.rdAll)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(613, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(249, 86)
        Me.GroupBox1.TabIndex = 3
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Destination"
        '
        'cbxReceiver
        '
        Me.cbxReceiver.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxReceiver.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxReceiver.FormattingEnabled = True
        Me.cbxReceiver.Location = New System.Drawing.Point(6, 48)
        Me.cbxReceiver.Name = "cbxReceiver"
        Me.cbxReceiver.Size = New System.Drawing.Size(226, 24)
        Me.cbxReceiver.TabIndex = 3
        '
        'rdIndi
        '
        Me.rdIndi.AutoSize = True
        Me.rdIndi.Location = New System.Drawing.Point(50, 22)
        Me.rdIndi.Name = "rdIndi"
        Me.rdIndi.Size = New System.Drawing.Size(158, 21)
        Me.rdIndi.TabIndex = 2
        Me.rdIndi.Text = "Select Destination"
        Me.rdIndi.UseVisualStyleBackColor = True
        '
        'rdAll
        '
        Me.rdAll.AutoSize = True
        Me.rdAll.Checked = True
        Me.rdAll.Location = New System.Drawing.Point(6, 22)
        Me.rdAll.Name = "rdAll"
        Me.rdAll.Size = New System.Drawing.Size(44, 21)
        Me.rdAll.TabIndex = 1
        Me.rdAll.TabStop = True
        Me.rdAll.Text = "All"
        Me.rdAll.UseVisualStyleBackColor = True
        '
        'gbArea
        '
        Me.gbArea.Controls.Add(Me.cbxArea)
        Me.gbArea.Controls.Add(Me.rdSelectCategArea)
        Me.gbArea.Controls.Add(Me.rdAllArea)
        Me.gbArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbArea.Location = New System.Drawing.Point(868, 8)
        Me.gbArea.Name = "gbArea"
        Me.gbArea.Size = New System.Drawing.Size(243, 86)
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
        Me.cbxArea.Location = New System.Drawing.Point(6, 48)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(226, 24)
        Me.cbxArea.TabIndex = 6
        '
        'rdSelectCategArea
        '
        Me.rdSelectCategArea.AutoSize = True
        Me.rdSelectCategArea.Location = New System.Drawing.Point(51, 21)
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
        Me.rdAllArea.Location = New System.Drawing.Point(6, 21)
        Me.rdAllArea.Name = "rdAllArea"
        Me.rdAllArea.Size = New System.Drawing.Size(44, 21)
        Me.rdAllArea.TabIndex = 4
        Me.rdAllArea.TabStop = True
        Me.rdAllArea.Text = "All"
        Me.rdAllArea.UseVisualStyleBackColor = True
        '
        'releaseRegister
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.gbArea)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.lvRelsListing)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "releaseRegister"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Release Register"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gbArea.ResumeLayout(False)
        Me.gbArea.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvRelsListing As ListView
    Friend WithEvents lblError As Label
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents tbxDocType As TextBox
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents lblNumber As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label3 As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents dtpfrom As DateTimePicker
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents cbxReceiver As ComboBox
    Friend WithEvents rdIndi As RadioButton
    Friend WithEvents rdAll As RadioButton
    Friend WithEvents gbArea As GroupBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents rdSelectCategArea As RadioButton
    Friend WithEvents rdAllArea As RadioButton
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
End Class
