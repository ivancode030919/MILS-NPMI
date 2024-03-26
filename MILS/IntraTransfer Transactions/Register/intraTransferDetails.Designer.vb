<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class intraTransferDetails
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.tbxEntryNumber = New System.Windows.Forms.TextBox()
        Me.tbxApprovedBy = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.tbxRemarks = New System.Windows.Forms.TextBox()
        Me.tbxUser = New System.Windows.Forms.TextBox()
        Me.dtpTransferDate = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.tbxTransferId = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.lvIntraTransferDetails = New System.Windows.Forms.ListView()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.tbxEntryNumber)
        Me.GroupBox1.Controls.Add(Me.tbxApprovedBy)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.tbxRemarks)
        Me.GroupBox1.Controls.Add(Me.tbxUser)
        Me.GroupBox1.Controls.Add(Me.dtpTransferDate)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.tbxTransferId)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Top
        Me.GroupBox1.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(5, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1191, 113)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Details"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(766, 26)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(113, 18)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Entry Number :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(777, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(102, 18)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Approved By :"
        '
        'tbxEntryNumber
        '
        Me.tbxEntryNumber.Enabled = False
        Me.tbxEntryNumber.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxEntryNumber.Location = New System.Drawing.Point(885, 21)
        Me.tbxEntryNumber.Name = "tbxEntryNumber"
        Me.tbxEntryNumber.Size = New System.Drawing.Size(300, 23)
        Me.tbxEntryNumber.TabIndex = 1
        '
        'tbxApprovedBy
        '
        Me.tbxApprovedBy.Enabled = False
        Me.tbxApprovedBy.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxApprovedBy.Location = New System.Drawing.Point(885, 83)
        Me.tbxApprovedBy.Name = "tbxApprovedBy"
        Me.tbxApprovedBy.Size = New System.Drawing.Size(300, 23)
        Me.tbxApprovedBy.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(801, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 18)
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Remarks :"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(95, 86)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(75, 18)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Encoder :"
        '
        'tbxRemarks
        '
        Me.tbxRemarks.Enabled = False
        Me.tbxRemarks.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxRemarks.Location = New System.Drawing.Point(885, 51)
        Me.tbxRemarks.Name = "tbxRemarks"
        Me.tbxRemarks.Size = New System.Drawing.Size(300, 23)
        Me.tbxRemarks.TabIndex = 5
        '
        'tbxUser
        '
        Me.tbxUser.Enabled = False
        Me.tbxUser.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxUser.Location = New System.Drawing.Point(176, 81)
        Me.tbxUser.Name = "tbxUser"
        Me.tbxUser.Size = New System.Drawing.Size(300, 23)
        Me.tbxUser.TabIndex = 12
        '
        'dtpTransferDate
        '
        Me.dtpTransferDate.Enabled = False
        Me.dtpTransferDate.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTransferDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpTransferDate.Location = New System.Drawing.Point(176, 51)
        Me.dtpTransferDate.Name = "dtpTransferDate"
        Me.dtpTransferDate.Size = New System.Drawing.Size(300, 23)
        Me.dtpTransferDate.TabIndex = 3
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(47, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(123, 18)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Document Date :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(24, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(146, 18)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Document Number :"
        '
        'tbxTransferId
        '
        Me.tbxTransferId.Enabled = False
        Me.tbxTransferId.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbxTransferId.Location = New System.Drawing.Point(176, 21)
        Me.tbxTransferId.Name = "tbxTransferId"
        Me.tbxTransferId.Size = New System.Drawing.Size(300, 23)
        Me.tbxTransferId.TabIndex = 2
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.lvIntraTransferDetails)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.GroupBox3.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(5, 124)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(1191, 453)
        Me.GroupBox3.TabIndex = 3
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Contents"
        '
        'lvIntraTransferDetails
        '
        Me.lvIntraTransferDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvIntraTransferDetails.Font = New System.Drawing.Font("Franklin Gothic Book", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvIntraTransferDetails.FullRowSelect = True
        Me.lvIntraTransferDetails.GridLines = True
        Me.lvIntraTransferDetails.HideSelection = False
        Me.lvIntraTransferDetails.Location = New System.Drawing.Point(3, 19)
        Me.lvIntraTransferDetails.Name = "lvIntraTransferDetails"
        Me.lvIntraTransferDetails.Size = New System.Drawing.Size(1185, 431)
        Me.lvIntraTransferDetails.TabIndex = 0
        Me.lvIntraTransferDetails.UseCompatibleStateImageBehavior = False
        '
        'intraTransferDetails
        '
        Me.Appearance.Options.UseFont = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 17.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1201, 582)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.Font = New System.Drawing.Font("Franklin Gothic Book", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IconOptions.ShowIcon = False
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "intraTransferDetails"
        Me.Padding = New System.Windows.Forms.Padding(5)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Details"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label5 As Label
    Friend WithEvents tbxEntryNumber As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents tbxRemarks As TextBox
    Friend WithEvents dtpTransferDate As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbxTransferId As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents lvIntraTransferDetails As ListView
    Friend WithEvents Label6 As Label
    Friend WithEvents tbxUser As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents tbxApprovedBy As TextBox
End Class
