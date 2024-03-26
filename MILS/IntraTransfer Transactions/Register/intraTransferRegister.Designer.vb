<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class intraTransferRegister
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.dtpfrom = New System.Windows.Forms.DateTimePicker()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tbxEntryNum = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxTransferId = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lvIntraTransfer = New System.Windows.Forms.ListView()
        Me.lblError = New System.Windows.Forms.Label()
        Me.gbArea = New System.Windows.Forms.GroupBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.rdSelectCategArea = New System.Windows.Forms.RadioButton()
        Me.rdAllArea = New System.Windows.Forms.RadioButton()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.gbArea.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.dtpTo)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.dtpfrom)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(8, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(292, 82)
        Me.GroupBox2.TabIndex = 19
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Date Range"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(30, 28)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 17)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "From :"
        '
        'dtpTo
        '
        Me.dtpTo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(90, 50)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(187, 23)
        Me.dtpTo.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(47, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 17)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "To :"
        '
        'dtpfrom
        '
        Me.dtpfrom.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpfrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpfrom.Location = New System.Drawing.Point(90, 22)
        Me.dtpfrom.Name = "dtpfrom"
        Me.dtpfrom.Size = New System.Drawing.Size(187, 23)
        Me.dtpfrom.TabIndex = 2
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.tbxEntryNum)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.tbxTransferId)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox1.Location = New System.Drawing.Point(306, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(415, 82)
        Me.GroupBox1.TabIndex = 18
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Details"
        '
        'tbxEntryNum
        '
        Me.tbxEntryNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tbxEntryNum.Location = New System.Drawing.Point(185, 45)
        Me.tbxEntryNum.Name = "tbxEntryNum"
        Me.tbxEntryNum.Size = New System.Drawing.Size(215, 23)
        Me.tbxEntryNum.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(62, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(117, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Entry Number :"
        '
        'tbxTransferId
        '
        Me.tbxTransferId.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tbxTransferId.Location = New System.Drawing.Point(185, 17)
        Me.tbxTransferId.Name = "tbxTransferId"
        Me.tbxTransferId.Size = New System.Drawing.Size(215, 23)
        Me.tbxTransferId.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Document Number :"
        '
        'lvIntraTransfer
        '
        Me.lvIntraTransfer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvIntraTransfer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.lvIntraTransfer.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvIntraTransfer.HideSelection = False
        Me.lvIntraTransfer.Location = New System.Drawing.Point(2, 2)
        Me.lvIntraTransfer.Name = "lvIntraTransfer"
        Me.lvIntraTransfer.Size = New System.Drawing.Size(1187, 482)
        Me.lvIntraTransfer.TabIndex = 4
        Me.lvIntraTransfer.UseCompatibleStateImageBehavior = False
        '
        'lblError
        '
        Me.lblError.AutoSize = True
        Me.lblError.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblError.ForeColor = System.Drawing.Color.Red
        Me.lblError.Location = New System.Drawing.Point(533, 186)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(16, 20)
        Me.lblError.TabIndex = 22
        Me.lblError.Text = "*"
        Me.lblError.Visible = False
        '
        'gbArea
        '
        Me.gbArea.Controls.Add(Me.cbxArea)
        Me.gbArea.Controls.Add(Me.rdSelectCategArea)
        Me.gbArea.Controls.Add(Me.rdAllArea)
        Me.gbArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.gbArea.Location = New System.Drawing.Point(727, 3)
        Me.gbArea.Name = "gbArea"
        Me.gbArea.Size = New System.Drawing.Size(252, 82)
        Me.gbArea.TabIndex = 27
        Me.gbArea.TabStop = False
        Me.gbArea.Tag = "4"
        Me.gbArea.Text = "Area"
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(6, 48)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(240, 24)
        Me.cbxArea.TabIndex = 6
        '
        'rdSelectCategArea
        '
        Me.rdSelectCategArea.AutoSize = True
        Me.rdSelectCategArea.Location = New System.Drawing.Point(75, 21)
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
        Me.rdAllArea.Location = New System.Drawing.Point(30, 21)
        Me.rdAllArea.Name = "rdAllArea"
        Me.rdAllArea.Size = New System.Drawing.Size(44, 21)
        Me.rdAllArea.TabIndex = 4
        Me.rdAllArea.TabStop = True
        Me.rdAllArea.Text = "All"
        Me.rdAllArea.UseVisualStyleBackColor = True
        '
        'PanelControl1
        '
        Me.PanelControl1.Controls.Add(Me.lblError)
        Me.PanelControl1.Controls.Add(Me.lvIntraTransfer)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelControl1.Location = New System.Drawing.Point(5, 91)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(1191, 486)
        Me.PanelControl1.TabIndex = 28
        '
        'intraTransferRegister
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.PanelControl1)
        Me.Controls.Add(Me.gbArea)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "intraTransferRegister"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Transfer Register"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.gbArea.ResumeLayout(False)
        Me.gbArea.PerformLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Label3 As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents Label4 As Label
    Friend WithEvents dtpfrom As DateTimePicker
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents lvIntraTransfer As ListView
    Friend WithEvents lblError As Label
    Friend WithEvents tbxTransferId As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxEntryNum As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents gbArea As GroupBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents rdSelectCategArea As RadioButton
    Friend WithEvents rdAllArea As RadioButton
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
