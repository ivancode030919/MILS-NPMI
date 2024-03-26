<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class documentTypes
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxFilter1 = New System.Windows.Forms.TextBox()
        Me.lvDocTypes = New System.Windows.Forms.ListView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxFilter2 = New System.Windows.Forms.TextBox()
        Me.rdSelectType = New System.Windows.Forms.RadioButton()
        Me.rdAllType = New System.Windows.Forms.RadioButton()
        Me.cbxType = New System.Windows.Forms.ComboBox()
        Me.rdSelectCategory = New System.Windows.Forms.RadioButton()
        Me.rdAllCateg = New System.Windows.Forms.RadioButton()
        Me.cbxCategory = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.btnDetails = New System.Windows.Forms.Button()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 16)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Document :"
        '
        'tbxFilter1
        '
        Me.tbxFilter1.Location = New System.Drawing.Point(108, 9)
        Me.tbxFilter1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter1.Name = "tbxFilter1"
        Me.tbxFilter1.Size = New System.Drawing.Size(255, 21)
        Me.tbxFilter1.TabIndex = 1
        '
        'lvDocTypes
        '
        Me.lvDocTypes.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lvDocTypes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvDocTypes.HideSelection = False
        Me.lvDocTypes.Location = New System.Drawing.Point(0, 0)
        Me.lvDocTypes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvDocTypes.Name = "lvDocTypes"
        Me.lvDocTypes.Size = New System.Drawing.Size(696, 446)
        Me.lvDocTypes.TabIndex = 5
        Me.lvDocTypes.UseCompatibleStateImageBehavior = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(5, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 16)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Description :"
        '
        'tbxFilter2
        '
        Me.tbxFilter2.Location = New System.Drawing.Point(108, 36)
        Me.tbxFilter2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter2.Name = "tbxFilter2"
        Me.tbxFilter2.Size = New System.Drawing.Size(255, 21)
        Me.tbxFilter2.TabIndex = 2
        '
        'rdSelectType
        '
        Me.rdSelectType.AutoSize = True
        Me.rdSelectType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.rdSelectType.Location = New System.Drawing.Point(140, 5)
        Me.rdSelectType.Name = "rdSelectType"
        Me.rdSelectType.Size = New System.Drawing.Size(110, 20)
        Me.rdSelectType.TabIndex = 22
        Me.rdSelectType.Text = "Select Type"
        Me.rdSelectType.UseVisualStyleBackColor = True
        '
        'rdAllType
        '
        Me.rdAllType.AutoSize = True
        Me.rdAllType.Checked = True
        Me.rdAllType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold)
        Me.rdAllType.Location = New System.Drawing.Point(72, 5)
        Me.rdAllType.Name = "rdAllType"
        Me.rdAllType.Size = New System.Drawing.Size(44, 20)
        Me.rdAllType.TabIndex = 21
        Me.rdAllType.TabStop = True
        Me.rdAllType.Text = "All"
        Me.rdAllType.UseVisualStyleBackColor = True
        '
        'cbxType
        '
        Me.cbxType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxType.FormattingEnabled = True
        Me.cbxType.Items.AddRange(New Object() {"Primary", "Secondary"})
        Me.cbxType.Location = New System.Drawing.Point(54, 29)
        Me.cbxType.Name = "cbxType"
        Me.cbxType.Size = New System.Drawing.Size(250, 21)
        Me.cbxType.TabIndex = 20
        '
        'rdSelectCategory
        '
        Me.rdSelectCategory.AutoSize = True
        Me.rdSelectCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdSelectCategory.Location = New System.Drawing.Point(191, 67)
        Me.rdSelectCategory.Name = "rdSelectCategory"
        Me.rdSelectCategory.Size = New System.Drawing.Size(137, 20)
        Me.rdSelectCategory.TabIndex = 16
        Me.rdSelectCategory.Text = "Select Category"
        Me.rdSelectCategory.UseVisualStyleBackColor = True
        '
        'rdAllCateg
        '
        Me.rdAllCateg.AutoSize = True
        Me.rdAllCateg.Checked = True
        Me.rdAllCateg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdAllCateg.Location = New System.Drawing.Point(115, 67)
        Me.rdAllCateg.Name = "rdAllCateg"
        Me.rdAllCateg.Size = New System.Drawing.Size(44, 20)
        Me.rdAllCateg.TabIndex = 15
        Me.rdAllCateg.TabStop = True
        Me.rdAllCateg.Text = "All"
        Me.rdAllCateg.UseVisualStyleBackColor = True
        '
        'cbxCategory
        '
        Me.cbxCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCategory.FormattingEnabled = True
        Me.cbxCategory.Items.AddRange(New Object() {"Receiving", "Releasing"})
        Me.cbxCategory.Location = New System.Drawing.Point(108, 91)
        Me.cbxCategory.Name = "cbxCategory"
        Me.cbxCategory.Size = New System.Drawing.Size(255, 21)
        Me.cbxCategory.TabIndex = 13
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.PanelControl2)
        Me.Panel1.Controls.Add(Me.PanelControl1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(5, 5)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(700, 291)
        Me.Panel1.TabIndex = 17
        '
        'PanelControl2
        '
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.btnAdd)
        Me.PanelControl2.Controls.Add(Me.rdAllType)
        Me.PanelControl2.Controls.Add(Me.cbxType)
        Me.PanelControl2.Controls.Add(Me.rdSelectType)
        Me.PanelControl2.Controls.Add(Me.Label4)
        Me.PanelControl2.Controls.Add(Me.btnDetails)
        Me.PanelControl2.Location = New System.Drawing.Point(381, 11)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(312, 120)
        Me.PanelControl2.TabIndex = 8
        '
        'btnAdd
        '
        Me.btnAdd.BackColor = System.Drawing.Color.Transparent
        Me.btnAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnAdd.Image = Global.MILS.My.Resources.Resources.plus
        Me.btnAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnAdd.Location = New System.Drawing.Point(54, 63)
        Me.btnAdd.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(105, 49)
        Me.btnAdd.TabIndex = 3
        Me.btnAdd.Text = "Add"
        Me.btnAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnAdd.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(4, 7)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(52, 16)
        Me.Label4.TabIndex = 23
        Me.Label4.Text = "Type :"
        '
        'btnDetails
        '
        Me.btnDetails.BackColor = System.Drawing.Color.Transparent
        Me.btnDetails.Enabled = False
        Me.btnDetails.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDetails.Image = Global.MILS.My.Resources.Resources._loop
        Me.btnDetails.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDetails.Location = New System.Drawing.Point(199, 64)
        Me.btnDetails.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.btnDetails.Name = "btnDetails"
        Me.btnDetails.Size = New System.Drawing.Size(105, 49)
        Me.btnDetails.TabIndex = 4
        Me.btnDetails.Text = "Update"
        Me.btnDetails.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDetails.UseVisualStyleBackColor = False
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.cbxCategory)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Controls.Add(Me.tbxFilter1)
        Me.PanelControl1.Controls.Add(Me.tbxFilter2)
        Me.PanelControl1.Controls.Add(Me.rdAllCateg)
        Me.PanelControl1.Controls.Add(Me.rdSelectCategory)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Location = New System.Drawing.Point(3, 11)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(372, 120)
        Me.PanelControl1.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(21, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 16)
        Me.Label3.TabIndex = 17
        Me.Label3.Text = "Category :"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.lblErr)
        Me.Panel2.Controls.Add(Me.lvDocTypes)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(5, 141)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(700, 450)
        Me.Panel2.TabIndex = 18
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(294, 198)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(109, 17)
        Me.lblErr.TabIndex = 6
        Me.lblErr.Text = "no data found"
        '
        'documentTypes
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(710, 596)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "documentTypes"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Document Type"
        Me.Panel1.ResumeLayout(False)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        Me.PanelControl2.PerformLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxFilter1 As TextBox
    Friend WithEvents btnDetails As Button
    Friend WithEvents btnAdd As Button
    Friend WithEvents lvDocTypes As ListView
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxFilter2 As TextBox
    Friend WithEvents rdSelectType As RadioButton
    Friend WithEvents rdAllType As RadioButton
    Friend WithEvents cbxType As ComboBox
    Friend WithEvents rdSelectCategory As RadioButton
    Friend WithEvents rdAllCateg As RadioButton
    Friend WithEvents cbxCategory As ComboBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents lblErr As Label
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
