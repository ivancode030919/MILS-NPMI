<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class IGLRSpecific
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
        Me.crvIGLR = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cbxLoc = New System.Windows.Forms.ComboBox()
        Me.cbxBatch = New System.Windows.Forms.ComboBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxProdDes = New System.Windows.Forms.TextBox()
        Me.tbxProdNum = New System.Windows.Forms.TextBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.dtpStop = New System.Windows.Forms.DateTimePicker()
        Me.dtpStart = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'crvIGLR
        '
        Me.crvIGLR.ActiveViewIndex = -1
        Me.crvIGLR.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.crvIGLR.Cursor = System.Windows.Forms.Cursors.Default
        Me.crvIGLR.Dock = System.Windows.Forms.DockStyle.Fill
        Me.crvIGLR.Location = New System.Drawing.Point(2, 2)
        Me.crvIGLR.Name = "crvIGLR"
        Me.crvIGLR.Size = New System.Drawing.Size(925, 627)
        Me.crvIGLR.TabIndex = 9
        Me.crvIGLR.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cbxLoc)
        Me.GroupBox4.Controls.Add(Me.cbxBatch)
        Me.GroupBox4.Controls.Add(Me.cbxArea)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.Label5)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.Location = New System.Drawing.Point(12, 223)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(332, 116)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Details"
        '
        'cbxLoc
        '
        Me.cbxLoc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxLoc.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxLoc.FormattingEnabled = True
        Me.cbxLoc.Location = New System.Drawing.Point(94, 83)
        Me.cbxLoc.Name = "cbxLoc"
        Me.cbxLoc.Size = New System.Drawing.Size(232, 25)
        Me.cbxLoc.TabIndex = 21
        '
        'cbxBatch
        '
        Me.cbxBatch.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxBatch.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxBatch.FormattingEnabled = True
        Me.cbxBatch.Location = New System.Drawing.Point(94, 52)
        Me.cbxBatch.Name = "cbxBatch"
        Me.cbxBatch.Size = New System.Drawing.Size(232, 25)
        Me.cbxBatch.TabIndex = 20
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(94, 21)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(232, 25)
        Me.cbxArea.TabIndex = 19
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label7.Location = New System.Drawing.Point(16, 86)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(76, 18)
        Me.Label7.TabIndex = 18
        Me.Label7.Text = "Location :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(35, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 18)
        Me.Label5.TabIndex = 16
        Me.Label5.Text = "Batch :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label6.Location = New System.Drawing.Point(43, 28)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(49, 18)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Area :"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.tbxProdDes)
        Me.GroupBox3.Controls.Add(Me.tbxProdNum)
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox3.Location = New System.Drawing.Point(12, 98)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(332, 119)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Product"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label4.Location = New System.Drawing.Point(8, 52)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 18)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Description :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(32, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 18)
        Me.Label1.TabIndex = 15
        Me.Label1.Text = "Number :"
        '
        'tbxProdDes
        '
        Me.tbxProdDes.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxProdDes.Location = New System.Drawing.Point(110, 49)
        Me.tbxProdDes.Multiline = True
        Me.tbxProdDes.Name = "tbxProdDes"
        Me.tbxProdDes.Size = New System.Drawing.Size(216, 59)
        Me.tbxProdDes.TabIndex = 4
        '
        'tbxProdNum
        '
        Me.tbxProdNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxProdNum.Location = New System.Drawing.Point(110, 21)
        Me.tbxProdNum.Name = "tbxProdNum"
        Me.tbxProdNum.Size = New System.Drawing.Size(216, 23)
        Me.tbxProdNum.TabIndex = 3
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.dtpStop)
        Me.GroupBox5.Controls.Add(Me.dtpStart)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox5.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(332, 80)
        Me.GroupBox5.TabIndex = 1
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Date Range"
        '
        'dtpStop
        '
        Me.dtpStop.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.dtpStop.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStop.Location = New System.Drawing.Point(85, 49)
        Me.dtpStop.Name = "dtpStop"
        Me.dtpStop.Size = New System.Drawing.Size(241, 23)
        Me.dtpStop.TabIndex = 2
        '
        'dtpStart
        '
        Me.dtpStart.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.dtpStart.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpStart.Location = New System.Drawing.Point(85, 21)
        Me.dtpStart.Name = "dtpStart"
        Me.dtpStart.Size = New System.Drawing.Size(241, 23)
        Me.dtpStart.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(27, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 18)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "From :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(45, 54)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 18)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "To :"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnClear)
        Me.GroupBox2.Controls.Add(Me.btnGenerate)
        Me.GroupBox2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 506)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(332, 127)
        Me.GroupBox2.TabIndex = 4
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Actions"
        '
        'btnClear
        '
        Me.btnClear.Font = New System.Drawing.Font("Franklin Gothic Book", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnClear.Location = New System.Drawing.Point(9, 21)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(317, 33)
        Me.btnClear.TabIndex = 6
        Me.btnClear.Text = "Clear Fields"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnGenerate
        '
        Me.btnGenerate.Font = New System.Drawing.Font("Franklin Gothic Book", 12.0!, System.Drawing.FontStyle.Bold)
        Me.btnGenerate.Location = New System.Drawing.Point(9, 60)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(317, 55)
        Me.btnGenerate.TabIndex = 7
        Me.btnGenerate.Text = "Generate Report"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.crvIGLR)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Right
        Me.PanelControl1.Location = New System.Drawing.Point(350, 5)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(929, 631)
        Me.PanelControl1.TabIndex = 39
        '
        'IGLRSpecific
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1284, 641)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox5)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "IGLRSpecific"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory Ledger [Specific]"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents crvIGLR As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents dtpStop As DateTimePicker
    Friend WithEvents dtpStart As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnClear As Button
    Friend WithEvents btnGenerate As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxProdNum As TextBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Label7 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents tbxProdDes As TextBox
    Friend WithEvents cbxLoc As ComboBox
    Friend WithEvents cbxBatch As ComboBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
