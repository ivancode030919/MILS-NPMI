<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class displayDocutypes
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
        Me.lvDoctypes = New System.Windows.Forms.ListView()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.rdSelectCategory = New System.Windows.Forms.RadioButton()
        Me.rdAllCateg = New System.Windows.Forms.RadioButton()
        Me.cbxCategory = New System.Windows.Forms.ComboBox()
        Me.rdSelectType = New System.Windows.Forms.RadioButton()
        Me.rdAllType = New System.Windows.Forms.RadioButton()
        Me.cbxType = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxDescrip = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxLocName = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.PanelControl2 = New DevExpress.XtraEditors.PanelControl()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvDoctypes
        '
        Me.lvDoctypes.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvDoctypes.HideSelection = False
        Me.lvDoctypes.Location = New System.Drawing.Point(0, 0)
        Me.lvDoctypes.Name = "lvDoctypes"
        Me.lvDoctypes.Size = New System.Drawing.Size(656, 311)
        Me.lvDoctypes.TabIndex = 4
        Me.lvDoctypes.UseCompatibleStateImageBehavior = False
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(17, 381)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(13, 16)
        Me.lblErr.TabIndex = 1
        Me.lblErr.Text = "*"
        '
        'rdSelectCategory
        '
        Me.rdSelectCategory.AutoSize = True
        Me.rdSelectCategory.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdSelectCategory.Location = New System.Drawing.Point(159, 8)
        Me.rdSelectCategory.Name = "rdSelectCategory"
        Me.rdSelectCategory.Size = New System.Drawing.Size(141, 21)
        Me.rdSelectCategory.TabIndex = 16
        Me.rdSelectCategory.Text = "Select Category"
        Me.rdSelectCategory.UseVisualStyleBackColor = True
        '
        'rdAllCateg
        '
        Me.rdAllCateg.AutoSize = True
        Me.rdAllCateg.Checked = True
        Me.rdAllCateg.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdAllCateg.Location = New System.Drawing.Point(95, 10)
        Me.rdAllCateg.Name = "rdAllCateg"
        Me.rdAllCateg.Size = New System.Drawing.Size(44, 21)
        Me.rdAllCateg.TabIndex = 15
        Me.rdAllCateg.TabStop = True
        Me.rdAllCateg.Text = "All"
        Me.rdAllCateg.UseVisualStyleBackColor = True
        '
        'cbxCategory
        '
        Me.cbxCategory.Dock = System.Windows.Forms.DockStyle.Top
        Me.cbxCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxCategory.FormattingEnabled = True
        Me.cbxCategory.Items.AddRange(New Object() {"Receiving", "Releasing"})
        Me.cbxCategory.Location = New System.Drawing.Point(0, 0)
        Me.cbxCategory.Name = "cbxCategory"
        Me.cbxCategory.Size = New System.Drawing.Size(201, 21)
        Me.cbxCategory.TabIndex = 13
        '
        'rdSelectType
        '
        Me.rdSelectType.AutoSize = True
        Me.rdSelectType.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdSelectType.Location = New System.Drawing.Point(204, 69)
        Me.rdSelectType.Name = "rdSelectType"
        Me.rdSelectType.Size = New System.Drawing.Size(112, 21)
        Me.rdSelectType.TabIndex = 22
        Me.rdSelectType.Text = "Select Type"
        Me.rdSelectType.UseVisualStyleBackColor = True
        '
        'rdAllType
        '
        Me.rdAllType.AutoSize = True
        Me.rdAllType.Checked = True
        Me.rdAllType.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rdAllType.Location = New System.Drawing.Point(111, 69)
        Me.rdAllType.Name = "rdAllType"
        Me.rdAllType.Size = New System.Drawing.Size(44, 21)
        Me.rdAllType.TabIndex = 21
        Me.rdAllType.TabStop = True
        Me.rdAllType.Text = "All"
        Me.rdAllType.UseVisualStyleBackColor = True
        '
        'cbxType
        '
        Me.cbxType.Dock = System.Windows.Forms.DockStyle.Top
        Me.cbxType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbxType.FormattingEnabled = True
        Me.cbxType.Items.AddRange(New Object() {"Primary", "Secondary"})
        Me.cbxType.Location = New System.Drawing.Point(0, 0)
        Me.cbxType.Name = "cbxType"
        Me.cbxType.Size = New System.Drawing.Size(203, 21)
        Me.cbxType.TabIndex = 20
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(5, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 17)
        Me.Label2.TabIndex = 10
        Me.Label2.Text = "Description :"
        '
        'tbxDescrip
        '
        Me.tbxDescrip.Location = New System.Drawing.Point(111, 40)
        Me.tbxDescrip.Name = "tbxDescrip"
        Me.tbxDescrip.Size = New System.Drawing.Size(205, 21)
        Me.tbxDescrip.TabIndex = 9
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Document :"
        '
        'tbxLocName
        '
        Me.tbxLocName.Location = New System.Drawing.Point(111, 8)
        Me.tbxLocName.Name = "tbxLocName"
        Me.tbxLocName.Size = New System.Drawing.Size(205, 21)
        Me.tbxLocName.TabIndex = 7
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.lvDoctypes)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(5, 149)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(660, 315)
        Me.Panel1.TabIndex = 5
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.PanelControl2)
        Me.Panel2.Controls.Add(Me.PanelControl1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(5, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(660, 141)
        Me.Panel2.TabIndex = 6
        '
        'PanelControl2
        '
        Me.PanelControl2.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl2.Controls.Add(Me.rdSelectCategory)
        Me.PanelControl2.Controls.Add(Me.rdAllCateg)
        Me.PanelControl2.Controls.Add(Me.Panel4)
        Me.PanelControl2.Controls.Add(Me.Label4)
        Me.PanelControl2.Location = New System.Drawing.Point(335, 3)
        Me.PanelControl2.Name = "PanelControl2"
        Me.PanelControl2.Size = New System.Drawing.Size(318, 129)
        Me.PanelControl2.TabIndex = 8
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel4.Controls.Add(Me.cbxCategory)
        Me.Panel4.Location = New System.Drawing.Point(95, 42)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(205, 28)
        Me.Panel4.TabIndex = 25
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(4, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(83, 17)
        Me.Label4.TabIndex = 26
        Me.Label4.Text = "Category :"
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.rdSelectType)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Controls.Add(Me.tbxLocName)
        Me.PanelControl1.Controls.Add(Me.tbxDescrip)
        Me.PanelControl1.Controls.Add(Me.rdAllType)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.Panel3)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Location = New System.Drawing.Point(3, 3)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(326, 129)
        Me.PanelControl1.TabIndex = 7
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.cbxType)
        Me.Panel3.Location = New System.Drawing.Point(111, 96)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(207, 26)
        Me.Panel3.TabIndex = 23
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(51, 71)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(54, 17)
        Me.Label3.TabIndex = 24
        Me.Label3.Text = "Type :"
        '
        'displayDocutypes
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(670, 469)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.lblErr)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "displayDocutypes"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Document Type"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.PanelControl2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl2.ResumeLayout(False)
        Me.PanelControl2.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lvDoctypes As ListView
    Friend WithEvents lblErr As Label
    Friend WithEvents rdSelectCategory As RadioButton
    Friend WithEvents rdAllCateg As RadioButton
    Friend WithEvents cbxCategory As ComboBox
    Friend WithEvents rdSelectType As RadioButton
    Friend WithEvents rdAllType As RadioButton
    Friend WithEvents cbxType As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxDescrip As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxLocName As TextBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Label4 As Label
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
    Friend WithEvents PanelControl2 As DevExpress.XtraEditors.PanelControl
End Class
