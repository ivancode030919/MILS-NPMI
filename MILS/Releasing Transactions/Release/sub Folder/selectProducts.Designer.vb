<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class selectProducts
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxFilter2 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxFilter1 = New System.Windows.Forms.TextBox()
        Me.lvProducts = New System.Windows.Forms.ListView()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(375, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 18)
        Me.Label2.TabIndex = 21
        Me.Label2.Text = "Description :"
        '
        'tbxFilter2
        '
        Me.tbxFilter2.Dock = System.Windows.Forms.DockStyle.Top
        Me.tbxFilter2.Location = New System.Drawing.Point(0, 0)
        Me.tbxFilter2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter2.Name = "tbxFilter2"
        Me.tbxFilter2.Size = New System.Drawing.Size(273, 22)
        Me.tbxFilter2.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(14, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(131, 18)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Product Number :"
        '
        'tbxFilter1
        '
        Me.tbxFilter1.Dock = System.Windows.Forms.DockStyle.Top
        Me.tbxFilter1.Location = New System.Drawing.Point(0, 0)
        Me.tbxFilter1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter1.Name = "tbxFilter1"
        Me.tbxFilter1.Size = New System.Drawing.Size(194, 22)
        Me.tbxFilter1.TabIndex = 1
        '
        'lvProducts
        '
        Me.lvProducts.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lvProducts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvProducts.HideSelection = False
        Me.lvProducts.Location = New System.Drawing.Point(0, 0)
        Me.lvProducts.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvProducts.Name = "lvProducts"
        Me.lvProducts.Size = New System.Drawing.Size(903, 318)
        Me.lvProducts.TabIndex = 3
        Me.lvProducts.UseCompatibleStateImageBehavior = False
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel2.Controls.Add(Me.tbxFilter1)
        Me.Panel2.Location = New System.Drawing.Point(149, 10)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(198, 24)
        Me.Panel2.TabIndex = 24
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.tbxFilter2)
        Me.Panel1.Location = New System.Drawing.Point(477, 10)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(277, 24)
        Me.Panel1.TabIndex = 9
        '
        'Panel3
        '
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.lvProducts)
        Me.Panel3.Location = New System.Drawing.Point(15, 41)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(907, 322)
        Me.Panel3.TabIndex = 9
        '
        'selectProducts
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(937, 377)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "selectProducts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "List of Goods"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label2 As Label
    Friend WithEvents tbxFilter2 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxFilter1 As TextBox
    Friend WithEvents lvProducts As ListView
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel3 As Panel
End Class
