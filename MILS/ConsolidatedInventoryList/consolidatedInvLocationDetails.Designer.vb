<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class consolidatedInvLocationDetails
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
        Me.Label5 = New System.Windows.Forms.Label()
        Me.tbxBatch = New System.Windows.Forms.TextBox()
        Me.locStartFrm = New System.Windows.Forms.Label()
        Me.tbxProdDes = New System.Windows.Forms.TextBox()
        Me.tbxLocStartFrom = New System.Windows.Forms.TextBox()
        Me.tbxProdNum = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rdSelectCateg = New System.Windows.Forms.RadioButton()
        Me.rdAll = New System.Windows.Forms.RadioButton()
        Me.cbxArea = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.lvConsoInv = New System.Windows.Forms.ListView()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label5.Location = New System.Drawing.Point(401, 54)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(57, 18)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "Batch :"
        '
        'tbxBatch
        '
        Me.tbxBatch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxBatch.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxBatch.Location = New System.Drawing.Point(464, 51)
        Me.tbxBatch.Name = "tbxBatch"
        Me.tbxBatch.Size = New System.Drawing.Size(164, 23)
        Me.tbxBatch.TabIndex = 13
        '
        'locStartFrm
        '
        Me.locStartFrm.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.locStartFrm.AutoSize = True
        Me.locStartFrm.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.locStartFrm.Location = New System.Drawing.Point(382, 25)
        Me.locStartFrm.Name = "locStartFrm"
        Me.locStartFrm.Size = New System.Drawing.Size(76, 18)
        Me.locStartFrm.TabIndex = 10
        Me.locStartFrm.Text = "Location :"
        '
        'tbxProdDes
        '
        Me.tbxProdDes.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxProdDes.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxProdDes.Location = New System.Drawing.Point(157, 51)
        Me.tbxProdDes.Name = "tbxProdDes"
        Me.tbxProdDes.Size = New System.Drawing.Size(191, 23)
        Me.tbxProdDes.TabIndex = 2
        '
        'tbxLocStartFrom
        '
        Me.tbxLocStartFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxLocStartFrom.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxLocStartFrom.Location = New System.Drawing.Point(464, 20)
        Me.tbxLocStartFrom.Name = "tbxLocStartFrom"
        Me.tbxLocStartFrom.Size = New System.Drawing.Size(164, 23)
        Me.tbxLocStartFrom.TabIndex = 9
        '
        'tbxProdNum
        '
        Me.tbxProdNum.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tbxProdNum.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.tbxProdNum.Location = New System.Drawing.Point(157, 22)
        Me.tbxProdNum.Name = "tbxProdNum"
        Me.tbxProdNum.Size = New System.Drawing.Size(191, 23)
        Me.tbxProdNum.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(20, 27)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Product Number :"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(55, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 18)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Description :"
        '
        'rdSelectCateg
        '
        Me.rdSelectCateg.AutoSize = True
        Me.rdSelectCateg.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.rdSelectCateg.Location = New System.Drawing.Point(54, 21)
        Me.rdSelectCateg.Name = "rdSelectCateg"
        Me.rdSelectCateg.Size = New System.Drawing.Size(104, 22)
        Me.rdSelectCateg.TabIndex = 7
        Me.rdSelectCateg.Text = "Select Area"
        Me.rdSelectCateg.UseVisualStyleBackColor = True
        '
        'rdAll
        '
        Me.rdAll.AutoSize = True
        Me.rdAll.Checked = True
        Me.rdAll.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.rdAll.Location = New System.Drawing.Point(9, 21)
        Me.rdAll.Name = "rdAll"
        Me.rdAll.Size = New System.Drawing.Size(43, 22)
        Me.rdAll.TabIndex = 6
        Me.rdAll.TabStop = True
        Me.rdAll.Text = "All"
        Me.rdAll.UseVisualStyleBackColor = True
        '
        'cbxArea
        '
        Me.cbxArea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxArea.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.cbxArea.FormattingEnabled = True
        Me.cbxArea.Location = New System.Drawing.Point(9, 51)
        Me.cbxArea.Name = "cbxArea"
        Me.cbxArea.Size = New System.Drawing.Size(186, 25)
        Me.cbxArea.TabIndex = 4
        Me.cbxArea.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lblErr)
        Me.GroupBox2.Controls.Add(Me.lvConsoInv)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox2.Location = New System.Drawing.Point(5, 94)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1191, 483)
        Me.GroupBox2.TabIndex = 3
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Details"
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(520, 203)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(16, 20)
        Me.lblErr.TabIndex = 19
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'lvConsoInv
        '
        Me.lvConsoInv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lvConsoInv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvConsoInv.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!)
        Me.lvConsoInv.HideSelection = False
        Me.lvConsoInv.Location = New System.Drawing.Point(3, 19)
        Me.lvConsoInv.Name = "lvConsoInv"
        Me.lvConsoInv.Size = New System.Drawing.Size(1185, 461)
        Me.lvConsoInv.TabIndex = 0
        Me.lvConsoInv.UseCompatibleStateImageBehavior = False
        Me.lvConsoInv.View = System.Windows.Forms.View.Details
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cbxArea)
        Me.GroupBox3.Controls.Add(Me.rdAll)
        Me.GroupBox3.Controls.Add(Me.rdSelectCateg)
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox3.Location = New System.Drawing.Point(645, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(204, 85)
        Me.GroupBox3.TabIndex = 16
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Area"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.tbxBatch)
        Me.GroupBox5.Controls.Add(Me.Label1)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.Label3)
        Me.GroupBox5.Controls.Add(Me.tbxLocStartFrom)
        Me.GroupBox5.Controls.Add(Me.tbxProdDes)
        Me.GroupBox5.Controls.Add(Me.locStartFrm)
        Me.GroupBox5.Controls.Add(Me.tbxProdNum)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.GroupBox5.Location = New System.Drawing.Point(5, 3)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(634, 85)
        Me.GroupBox5.TabIndex = 18
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Search"
        '
        'consolidatedInvLocationDetails
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox3)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "consolidatedInvLocationDetails"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory List by Location"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label5 As Label
    Friend WithEvents tbxBatch As TextBox
    Friend WithEvents locStartFrm As Label
    Friend WithEvents tbxProdDes As TextBox
    Friend WithEvents tbxLocStartFrom As TextBox
    Friend WithEvents tbxProdNum As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents rdSelectCateg As RadioButton
    Friend WithEvents rdAll As RadioButton
    Friend WithEvents cbxArea As ComboBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents lvConsoInv As ListView
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents lblErr As Label
End Class
