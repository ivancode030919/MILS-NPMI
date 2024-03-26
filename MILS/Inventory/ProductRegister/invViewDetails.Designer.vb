<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class invViewDetails
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
        Me.lblErr = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxBatch = New System.Windows.Forms.TextBox()
        Me.lvProdDes = New System.Windows.Forms.ListView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tbxLoc = New System.Windows.Forms.TextBox()
        Me.tbxTotalQty = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.PanelControl1 = New DevExpress.XtraEditors.PanelControl()
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.PanelControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblErr
        '
        Me.lblErr.AutoSize = True
        Me.lblErr.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lblErr.ForeColor = System.Drawing.Color.Red
        Me.lblErr.Location = New System.Drawing.Point(438, 165)
        Me.lblErr.Name = "lblErr"
        Me.lblErr.Size = New System.Drawing.Size(17, 18)
        Me.lblErr.TabIndex = 11
        Me.lblErr.Text = "*"
        Me.lblErr.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(12, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 18)
        Me.Label1.TabIndex = 9
        Me.Label1.Text = "Batch Code:"
        '
        'tbxBatch
        '
        Me.tbxBatch.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tbxBatch.Location = New System.Drawing.Point(109, 13)
        Me.tbxBatch.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.tbxBatch.Name = "tbxBatch"
        Me.tbxBatch.Size = New System.Drawing.Size(297, 23)
        Me.tbxBatch.TabIndex = 6
        '
        'lvProdDes
        '
        Me.lvProdDes.BackColor = System.Drawing.SystemColors.InactiveBorder
        Me.lvProdDes.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.lvProdDes.HideSelection = False
        Me.lvProdDes.Location = New System.Drawing.Point(12, 47)
        Me.lvProdDes.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.lvProdDes.Name = "lvProdDes"
        Me.lvProdDes.Size = New System.Drawing.Size(960, 285)
        Me.lvProdDes.TabIndex = 10
        Me.lvProdDes.UseCompatibleStateImageBehavior = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label2.Location = New System.Drawing.Point(426, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(71, 18)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "Location:"
        '
        'tbxLoc
        '
        Me.tbxLoc.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tbxLoc.Location = New System.Drawing.Point(503, 13)
        Me.tbxLoc.Margin = New System.Windows.Forms.Padding(3, 5, 3, 5)
        Me.tbxLoc.Name = "tbxLoc"
        Me.tbxLoc.Size = New System.Drawing.Size(297, 23)
        Me.tbxLoc.TabIndex = 12
        '
        'tbxTotalQty
        '
        Me.tbxTotalQty.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tbxTotalQty.Location = New System.Drawing.Point(722, 340)
        Me.tbxTotalQty.Name = "tbxTotalQty"
        Me.tbxTotalQty.ReadOnly = True
        Me.tbxTotalQty.Size = New System.Drawing.Size(250, 23)
        Me.tbxTotalQty.TabIndex = 29
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Label3.Location = New System.Drawing.Point(608, 345)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(108, 18)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Total Quantity:"
        '
        'PanelControl1
        '
        Me.PanelControl1.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder
        Me.PanelControl1.Controls.Add(Me.Label1)
        Me.PanelControl1.Controls.Add(Me.tbxTotalQty)
        Me.PanelControl1.Controls.Add(Me.lvProdDes)
        Me.PanelControl1.Controls.Add(Me.Label3)
        Me.PanelControl1.Controls.Add(Me.tbxBatch)
        Me.PanelControl1.Controls.Add(Me.Label2)
        Me.PanelControl1.Controls.Add(Me.lblErr)
        Me.PanelControl1.Controls.Add(Me.tbxLoc)
        Me.PanelControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelControl1.Location = New System.Drawing.Point(0, 0)
        Me.PanelControl1.Name = "PanelControl1"
        Me.PanelControl1.Size = New System.Drawing.Size(986, 378)
        Me.PanelControl1.TabIndex = 30
        '
        'invViewDetails
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(986, 378)
        Me.Controls.Add(Me.PanelControl1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "invViewDetails"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Details"
        CType(Me.PanelControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.PanelControl1.ResumeLayout(False)
        Me.PanelControl1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents lblErr As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxBatch As TextBox
    Friend WithEvents lvProdDes As ListView
    Friend WithEvents Label2 As Label
    Friend WithEvents tbxLoc As TextBox
    Friend WithEvents tbxTotalQty As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents PanelControl1 As DevExpress.XtraEditors.PanelControl
End Class
