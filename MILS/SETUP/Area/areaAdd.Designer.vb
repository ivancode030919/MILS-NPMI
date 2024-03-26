<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class areaAdd
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        Me.tbxAreaName = New System.Windows.Forms.TextBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbxAreaDes = New System.Windows.Forms.TextBox()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.rdInternal = New System.Windows.Forms.RadioButton()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.rdExternal = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.PanelControl1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(384, 196)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Details"
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.tbxAreaName)
        Me.PanelControl1.Controls.Add(Me.lblErr)
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Controls.Add(Me.Label4)
        Me.PanelControl1.Controls.Add(Me.tbxAreaDes)
        Me.PanelControl1.Controls.Add(Me.btnCancel)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.btnAdd)
        Me.PanelControl1.Controls.Add(Me.rdInternal)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Controls.Add(Me.rdExternal)
        Me.PanelControl1.Location = New System.Drawing.Point(12, 21)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(359, 165)
        Me.PanelControl1.TabIndex = 20
        '
        'tbxAreaName
        '
        Me.tbxAreaName.Location = New System.Drawing.Point(102, 6)
        Me.tbxAreaName.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxAreaName.Name = "tbxAreaName"
        Me.tbxAreaName.Size = New System.Drawing.Size(249, 21)
        Me.tbxAreaName.TabIndex = 1
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(7, 137)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(13, 13)
        Me.lblErr.TabIndex = 19
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(23, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Area Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.ForeColor = System.Drawing.Color.Red
        Me.Label4.Location = New System.Drawing.Point(1, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(13, 13)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "*"
        '
        'tbxAreaDes
        '
        Me.tbxAreaDes.Location = New System.Drawing.Point(102, 36)
        Me.tbxAreaDes.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxAreaDes.Name = "tbxAreaDes"
        Me.tbxAreaDes.Size = New System.Drawing.Size(249, 21)
        Me.tbxAreaDes.TabIndex = 2
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(276, 133)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(75, 25)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(23, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(64, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Description:"
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(195, 133)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(75, 25)
        Me.btnAdd.TabIndex = 5
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'rdInternal
        '
        Me.rdInternal.AutoSize = True
        Me.rdInternal.Checked = True
        Me.rdInternal.Location = New System.Drawing.Point(102, 65)
        Me.rdInternal.Name = "rdInternal"
        Me.rdInternal.Size = New System.Drawing.Size(63, 17)
        Me.rdInternal.TabIndex = 3
        Me.rdInternal.TabStop = True
        Me.rdInternal.Text = "Internal"
        Me.rdInternal.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(23, 67)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Category:"
        '
        'rdExternal
        '
        Me.rdExternal.AutoSize = True
        Me.rdExternal.Location = New System.Drawing.Point(102, 92)
        Me.rdExternal.Name = "rdExternal"
        Me.rdExternal.Size = New System.Drawing.Size(65, 17)
        Me.rdExternal.TabIndex = 4
        Me.rdExternal.Text = "External"
        Me.rdExternal.UseVisualStyleBackColor = True
        '
        'areaAdd
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 196)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "areaAdd"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Add"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label4 As Label
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnAdd As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents rdExternal As RadioButton
    Friend WithEvents rdInternal As RadioButton
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxAreaDes As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxAreaName As TextBox
    Friend WithEvents lblErr As Label
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
