<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class itemListSmmary
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
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.cbxProdDes = New System.Windows.Forms.CheckedListBox()
        Me.tbxProdDes = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbxProdId = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.clbAreas = New System.Windows.Forms.CheckedListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnGenerate = New System.Windows.Forms.Button()
        Me.crvIGLR = New CrystalDecisions.Windows.Forms.CrystalReportViewer()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.cbxProdDes)
        Me.GroupBox4.Controls.Add(Me.tbxProdDes)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.Label4)
        Me.GroupBox4.Controls.Add(Me.tbxProdId)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(338, 206)
        Me.GroupBox4.TabIndex = 27
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Product"
        '
        'cbxProdDes
        '
        Me.cbxProdDes.FormattingEnabled = True
        Me.cbxProdDes.Location = New System.Drawing.Point(6, 77)
        Me.cbxProdDes.Name = "cbxProdDes"
        Me.cbxProdDes.Size = New System.Drawing.Size(326, 123)
        Me.cbxProdDes.TabIndex = 5
        '
        'tbxProdDes
        '
        Me.tbxProdDes.Location = New System.Drawing.Point(87, 49)
        Me.tbxProdDes.Name = "tbxProdDes"
        Me.tbxProdDes.Size = New System.Drawing.Size(243, 22)
        Me.tbxProdDes.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 52)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(79, 16)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Description:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 24)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(59, 16)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "Number:"
        '
        'tbxProdId
        '
        Me.tbxProdId.Location = New System.Drawing.Point(87, 21)
        Me.tbxProdId.Name = "tbxProdId"
        Me.tbxProdId.Size = New System.Drawing.Size(243, 22)
        Me.tbxProdId.TabIndex = 3
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.clbAreas)
        Me.GroupBox3.Location = New System.Drawing.Point(8, 224)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(338, 179)
        Me.GroupBox3.TabIndex = 26
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Area"
        '
        'clbAreas
        '
        Me.clbAreas.Dock = System.Windows.Forms.DockStyle.Fill
        Me.clbAreas.FormattingEnabled = True
        Me.clbAreas.Location = New System.Drawing.Point(3, 18)
        Me.clbAreas.Name = "clbAreas"
        Me.clbAreas.Size = New System.Drawing.Size(332, 158)
        Me.clbAreas.TabIndex = 0
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnClear)
        Me.GroupBox2.Controls.Add(Me.btnGenerate)
        Me.GroupBox2.Location = New System.Drawing.Point(7, 465)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(338, 127)
        Me.GroupBox2.TabIndex = 25
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Actions"
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(9, 21)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(321, 44)
        Me.btnClear.TabIndex = 6
        Me.btnClear.Text = "Clear Fields"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'btnGenerate
        '
        Me.btnGenerate.Location = New System.Drawing.Point(9, 71)
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.Size = New System.Drawing.Size(321, 44)
        Me.btnGenerate.TabIndex = 7
        Me.btnGenerate.Text = "Generate Report"
        Me.btnGenerate.UseVisualStyleBackColor = True
        '
        'crvIGLR
        '
        Me.crvIGLR.ActiveViewIndex = -1
        Me.crvIGLR.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.crvIGLR.Cursor = System.Windows.Forms.Cursors.Default
        Me.crvIGLR.Dock = System.Windows.Forms.DockStyle.Right
        Me.crvIGLR.Location = New System.Drawing.Point(356, 0)
        Me.crvIGLR.Name = "crvIGLR"
        Me.crvIGLR.Size = New System.Drawing.Size(928, 641)
        Me.crvIGLR.TabIndex = 28
        Me.crvIGLR.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.GroupBox2)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(350, 641)
        Me.GroupBox1.TabIndex = 29
        Me.GroupBox1.TabStop = False
        '
        'itemListSmmary
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1284, 641)
        Me.Controls.Add(Me.crvIGLR)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "itemListSmmary"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inventory List Summary"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents cbxProdDes As CheckedListBox
    Friend WithEvents tbxProdDes As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents tbxProdId As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents clbAreas As CheckedListBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents btnClear As Button
    Friend WithEvents btnGenerate As Button
    Friend WithEvents crvIGLR As CrystalDecisions.Windows.Forms.CrystalReportViewer
    Friend WithEvents GroupBox1 As GroupBox
End Class
