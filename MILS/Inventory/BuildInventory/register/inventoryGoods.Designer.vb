<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class inventoryGoods
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
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.tbxDocNum = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbxEntryNum = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.gbArea = New System.Windows.Forms.GroupBox()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.rdSelectCategArea = New System.Windows.Forms.RadioButton()
        Me.rdAllArea = New System.Windows.Forms.RadioButton()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.lvInventory = New System.Windows.Forms.ListView()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.gbArea.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.tbxDocNum)
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.tbxEntryNum)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox5.Location = New System.Drawing.Point(276, 8)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(451, 87)
        Me.GroupBox5.TabIndex = 2
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Details"
        '
        'tbxDocNum
        '
        Me.tbxDocNum.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxDocNum.Location = New System.Drawing.Point(185, 21)
        Me.tbxDocNum.Name = "tbxDocNum"
        Me.tbxDocNum.Size = New System.Drawing.Size(250, 22)
        Me.tbxDocNum.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(28, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(146, 18)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Document Number :"
        '
        'tbxEntryNum
        '
        Me.tbxEntryNum.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEntryNum.Location = New System.Drawing.Point(185, 50)
        Me.tbxEntryNum.Name = "tbxEntryNum"
        Me.tbxEntryNum.Size = New System.Drawing.Size(250, 22)
        Me.tbxEntryNum.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(62, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(113, 18)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Entry Number :"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.dtpFrom)
        Me.GroupBox4.Controls.Add(Me.Label2)
        Me.GroupBox4.Controls.Add(Me.dtpTo)
        Me.GroupBox4.Controls.Add(Me.Label3)
        Me.GroupBox4.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox4.Location = New System.Drawing.Point(5, 8)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(265, 87)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Date Range"
        '
        'dtpFrom
        '
        Me.dtpFrom.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpFrom.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpFrom.Location = New System.Drawing.Point(75, 22)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(170, 22)
        Me.dtpFrom.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(17, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(52, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "From :"
        '
        'dtpTo
        '
        Me.dtpTo.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTo.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTo.Location = New System.Drawing.Point(75, 50)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(170, 22)
        Me.dtpTo.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(35, 51)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "To :"
        '
        'gbArea
        '
        Me.gbArea.Controls.Add(Me.cbxArea)
        Me.gbArea.Controls.Add(Me.rdSelectCategArea)
        Me.gbArea.Controls.Add(Me.rdAllArea)
        Me.gbArea.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.gbArea.Location = New System.Drawing.Point(733, 8)
        Me.gbArea.Name = "gbArea"
        Me.gbArea.Size = New System.Drawing.Size(262, 87)
        Me.gbArea.TabIndex = 3
        Me.gbArea.TabStop = False
        Me.gbArea.Tag = "4"
        Me.gbArea.Text = "Area"
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(6, 48)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(250, 25)
        Me.cbxArea.TabIndex = 6
        '
        'rdSelectCategArea
        '
        Me.rdSelectCategArea.AutoSize = True
        Me.rdSelectCategArea.Location = New System.Drawing.Point(83, 21)
        Me.rdSelectCategArea.Name = "rdSelectCategArea"
        Me.rdSelectCategArea.Size = New System.Drawing.Size(104, 22)
        Me.rdSelectCategArea.TabIndex = 5
        Me.rdSelectCategArea.Text = "Select Area"
        Me.rdSelectCategArea.UseVisualStyleBackColor = True
        '
        'rdAllArea
        '
        Me.rdAllArea.AutoSize = True
        Me.rdAllArea.Checked = True
        Me.rdAllArea.Location = New System.Drawing.Point(38, 21)
        Me.rdAllArea.Name = "rdAllArea"
        Me.rdAllArea.Size = New System.Drawing.Size(43, 22)
        Me.rdAllArea.TabIndex = 4
        Me.rdAllArea.TabStop = True
        Me.rdAllArea.Text = "All"
        Me.rdAllArea.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lblErr)
        Me.GroupBox3.Controls.Add(Me.lvInventory)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(5, 101)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1191, 476)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Contents"
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(526, 181)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(20, 25)
        Me.lblErr.TabIndex = 4
        Me.lblErr.Text = "*"
        '
        'lvInventory
        '
        Me.lvInventory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvInventory.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvInventory.HideSelection = False
        Me.lvInventory.Location = New System.Drawing.Point(3, 19)
        Me.lvInventory.Name = "lvInventory"
        Me.lvInventory.Size = New System.Drawing.Size(1185, 454)
        Me.lvInventory.TabIndex = 0
        Me.lvInventory.UseCompatibleStateImageBehavior = False
        '
        'inventoryGoods
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.gbArea)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox4)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "inventoryGoods"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Build Inventory Register"
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.gbArea.ResumeLayout(False)
        Me.gbArea.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label3 As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents dtpFrom As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxEntryNum As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lvInventory As ListView
    Friend WithEvents tbxDocNum As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents gbArea As GroupBox
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents rdSelectCategArea As RadioButton
    Friend WithEvents rdAllArea As RadioButton
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents lblErr As Label
End Class
