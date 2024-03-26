<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class adjustInventoryRegister
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
        Me.cbxAreas = New System.Windows.Forms.ComboBox()
        Me.rdSelectCateg = New System.Windows.Forms.RadioButton()
        Me.rdAll = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxEntryNum = New System.Windows.Forms.TextBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lvInventoryAdjust = New System.Windows.Forms.ListView()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.SuspendLayout()
        '
        'cbxAreas
        '
        Me.cbxAreas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxAreas.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxAreas.FormattingEnabled = True
        Me.cbxAreas.Location = New System.Drawing.Point(6, 56)
        Me.cbxAreas.Name = "cbxAreas"
        Me.cbxAreas.Size = New System.Drawing.Size(197, 25)
        Me.cbxAreas.TabIndex = 6
        '
        'rdSelectCateg
        '
        Me.rdSelectCateg.AutoSize = True
        Me.rdSelectCateg.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.rdSelectCateg.Location = New System.Drawing.Point(89, 25)
        Me.rdSelectCateg.Name = "rdSelectCateg"
        Me.rdSelectCateg.Size = New System.Drawing.Size(104, 22)
        Me.rdSelectCateg.TabIndex = 5
        Me.rdSelectCateg.TabStop = True
        Me.rdSelectCateg.Text = "Select Area"
        Me.rdSelectCateg.UseVisualStyleBackColor = True
        '
        'rdAll
        '
        Me.rdAll.AutoSize = True
        Me.rdAll.Checked = True
        Me.rdAll.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.rdAll.Location = New System.Drawing.Point(23, 25)
        Me.rdAll.Name = "rdAll"
        Me.rdAll.Size = New System.Drawing.Size(43, 22)
        Me.rdAll.TabIndex = 4
        Me.rdAll.TabStop = True
        Me.rdAll.Text = "All"
        Me.rdAll.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(36, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Date To :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(16, 36)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 18)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Date From :"
        '
        'dtpTo
        '
        Me.dtpTo.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTo.Location = New System.Drawing.Point(115, 59)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(168, 22)
        Me.dtpTo.TabIndex = 2
        Me.dtpTo.Value = New Date(2023, 5, 26, 0, 0, 0, 0)
        '
        'dtpFrom
        '
        Me.dtpFrom.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpFrom.Location = New System.Drawing.Point(115, 31)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(168, 22)
        Me.dtpFrom.TabIndex = 2
        Me.dtpFrom.Value = New Date(2023, 5, 26, 0, 0, 0, 0)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(17, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Entry Number :"
        '
        'tbxEntryNum
        '
        Me.tbxEntryNum.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEntryNum.Location = New System.Drawing.Point(140, 36)
        Me.tbxEntryNum.Name = "tbxEntryNum"
        Me.tbxEntryNum.Size = New System.Drawing.Size(212, 22)
        Me.tbxEntryNum.TabIndex = 3
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold)
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(535, 190)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(16, 20)
        Me.lblErr.TabIndex = 1
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblErr)
        Me.GroupBox3.Controls.Add(Me.lvInventoryAdjust)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox3.Location = New System.Drawing.Point(5, 102)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1191, 475)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Content"
        '
        'lvInventoryAdjust
        '
        Me.lvInventoryAdjust.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvInventoryAdjust.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvInventoryAdjust.HideSelection = False
        Me.lvInventoryAdjust.Location = New System.Drawing.Point(3, 19)
        Me.lvInventoryAdjust.Name = "lvInventoryAdjust"
        Me.lvInventoryAdjust.Size = New System.Drawing.Size(1185, 453)
        Me.lvInventoryAdjust.TabIndex = 7
        Me.lvInventoryAdjust.UseCompatibleStateImageBehavior = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.dtpTo)
        Me.GroupBox4.Controls.Add(Me.dtpFrom)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(293, 96)
        Me.GroupBox4.TabIndex = 30
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Date Range"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.cbxAreas)
        Me.GroupBox1.Controls.Add(Me.rdAll)
        Me.GroupBox1.Controls.Add(Me.rdSelectCateg)
        Me.GroupBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.Location = New System.Drawing.Point(668, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(209, 96)
        Me.GroupBox1.TabIndex = 31
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Area"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label1)
        Me.GroupBox6.Controls.Add(Me.tbxEntryNum)
        Me.GroupBox6.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox6.Location = New System.Drawing.Point(304, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(358, 96)
        Me.GroupBox6.TabIndex = 33
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Search"
        '
        'adjustInventoryRegister
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "adjustInventoryRegister"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Count Register"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lvInventoryAdjust As ListView
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents dtpFrom As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxEntryNum As TextBox
    Friend WithEvents rdSelectCateg As RadioButton
    Friend WithEvents rdAll As RadioButton
    Friend WithEvents cbxAreas As ComboBox
    Friend WithEvents lblErr As Label
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox6 As GroupBox
End Class
