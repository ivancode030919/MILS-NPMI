<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class selectProdforReserveRelease
    Inherits System.Windows.Forms.Form

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
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxFilter2 = New System.Windows.Forms.TextBox()
        Me.lblErr = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxFilter1 = New System.Windows.Forms.TextBox()
        Me.lvProducts = New System.Windows.Forms.ListView()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(847, 368)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 30)
        Me.Button2.TabIndex = 31
        Me.Button2.Text = "Ok"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(766, 368)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 30)
        Me.Button1.TabIndex = 30
        Me.Button1.Text = "Cancel"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(380, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(73, 17)
        Me.Label2.TabIndex = 29
        Me.Label2.Text = "Description:"
        '
        'tbxFilter2
        '
        Me.tbxFilter2.Location = New System.Drawing.Point(459, 12)
        Me.tbxFilter2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter2.Name = "tbxFilter2"
        Me.tbxFilter2.Size = New System.Drawing.Size(255, 22)
        Me.tbxFilter2.TabIndex = 25
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(12, 375)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(16, 17)
        Me.lblErr.TabIndex = 28
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 15)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(101, 17)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Product Number:"
        '
        'tbxFilter1
        '
        Me.tbxFilter1.Location = New System.Drawing.Point(119, 12)
        Me.tbxFilter1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.tbxFilter1.Name = "tbxFilter1"
        Me.tbxFilter1.Size = New System.Drawing.Size(255, 22)
        Me.tbxFilter1.TabIndex = 24
        '
        'lvProducts
        '
        Me.lvProducts.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lvProducts.Location = New System.Drawing.Point(14, 42)
        Me.lvProducts.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.lvProducts.Name = "lvProducts"
        Me.lvProducts.Size = New System.Drawing.Size(908, 319)
        Me.lvProducts.TabIndex = 26
        Me.lvProducts.UseCompatibleStateImageBehavior = False
        '
        'selectProdforReserveRelease
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(934, 411)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.tbxFilter2)
        Me.Controls.Add(Me.lblErr)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.tbxFilter1)
        Me.Controls.Add(Me.lvProducts)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "selectProdforReserveRelease"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Select Product"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxFilter2 As TextBox
    Friend WithEvents lblErr As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxFilter1 As TextBox
    Friend WithEvents lvProducts As ListView
End Class
