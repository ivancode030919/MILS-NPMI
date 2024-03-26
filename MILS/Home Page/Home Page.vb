Imports System.Runtime.Remoting.Contexts
Imports System.Security
Imports System.Net
Public Class Home_Page
    Public userId As String
    Public roleId As String
    Public areaId As String
    Public empname As String
    Public branch As String
    Public notif1 As Boolean = False
    Private q As New qry
    Public alreadyClose As Boolean = False

    Private Sub Home_Page_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label3.Text = " IP Address: " & GetIPv4Address() + "             " + (Date.Now.ToString("MM/dd/yyyy") + "   " + " Version 3.0")
        hideUnnecessaryModules()
        Label1.Text = "User: " & empname & "            " + branch
        'MDIBGColor() -- pang ilis back ground sa MDIcontainer from gray to white
        'Label2.Text = branch



        If roleId = 20020 Or 20016 Then
            Timer2.Enabled = True
        Else
            Timer2.Enabled = False
        End If



    End Sub

    Function GetIPv4Address() As String
        Dim hostName As String = Dns.GetHostName()
        Dim ipAddresses As IPAddress() = Dns.GetHostEntry(hostName).AddressList

        Dim ipv4Address As String = ""
        For Each ipAddress As IPAddress In ipAddresses
            If ipAddress.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                ipv4Address = ipAddress.ToString()
                Exit For ' Only retrieve the first IPv4 address
            End If
        Next

        Return ipv4Address
    End Function

    Sub hideUnnecessaryModules()
        'Sales Crew ID Number
        'roleId = "20017" change to "20019"
        If roleId = "20019" Then
            InventoryListToolStripMenuItem.Visible = True
            ReserveToolStripMenuItem.Visible = False

            Exit Sub

            'Encoder ID number
            'roleId = "20019" change to "20018"
        ElseIf roleId = "20018" Then
            ListToolStripMenuItem.Visible = True
            ReceiveToolStripMenuItem.Visible = True
            ReleaseToolStripMenuItem.Visible = True
            IntraTransferToolStripMenuItem.Visible = True
            InventoryToolStripMenuItem1.Visible = True
            LocationToolStripMenuItem.Visible = True
            ReserveToolStripMenuItem.Visible = False
            ReasonCodeToolStripMenuItem.Visible = False
            'InventoryListToolStripMenuItem.Visible = True    
            Exit Sub

            'Supervisor ID Number
            'roleId = "20016" change to "20017"
        ElseIf roleId = "20017" Then
            ListToolStripMenuItem.Visible = True
            ReceiveToolStripMenuItem.Visible = True
            ReleaseToolStripMenuItem.Visible = True
            IntraTransferToolStripMenuItem.Visible = True
            InventoryListToolStripMenuItem.Visible = True
            InventoryToolStripMenuItem1.Visible = True
            LocationToolStripMenuItem.Visible = True
            ReserveToolStripMenuItem.Visible = False
            ReasonCodeToolStripMenuItem.Visible = False
            Exit Sub

            'Manager ID number
            'roleId = "20018" change to 20016
        ElseIf roleId = "20016" Then

            ListToolStripMenuItem.Visible = True
            ReceiveGoodsToolStripMenuItem.Visible = False
            ReleaseGoodsToolStripMenuItem.Visible = False
            RolesToolStripMenuItem.Visible = False
            TransferGoodsToolStripMenuItem1.Visible = False
            ReceiveToolStripMenuItem.Visible = True
            ReleaseToolStripMenuItem.Visible = True
            IntraTransferToolStripMenuItem.Visible = True
            InventoryToolStripMenuItem1.Visible = True
            ReportsToolStripMenuItem.Visible = True
            SetupToolStripMenuItem.Visible = True
            InventoryListToolStripMenuItem.Visible = True
            LocationToolStripMenuItem.Visible = True
            ReserveToolStripMenuItem.Visible = False
            ReasonCodeToolStripMenuItem.Visible = False
            NewToolStripMenuItem1.Visible = False
            NewToolStripMenuItem.Visible = False
            NotificationsToolStripMenuItem.Visible = True
            Exit Sub

            'Super User/ADMIN ID number
        ElseIf roleId = "20020" Then
            'EditorToolStripMenuItem.Visible = True
            ListToolStripMenuItem.Visible = True
            ReceiveToolStripMenuItem.Visible = True
            ReleaseToolStripMenuItem.Visible = True
            IntraTransferToolStripMenuItem.Visible = True
            InventoryToolStripMenuItem1.Visible = True
            ReportsToolStripMenuItem.Visible = True
            SetupToolStripMenuItem.Visible = True
            InventoryListToolStripMenuItem.Visible = True
            LocationToolStripMenuItem.Visible = True
            ReserveToolStripMenuItem.Visible = False
            NotificationsToolStripMenuItem.Visible = True
            ReasonCodeToolStripMenuItem.Visible = True
            OrderPointToolStripMenuItem.Visible = True
            EditorToolStripMenuItem1.Visible = True
            Exit Sub


        ElseIf roleId = "20021" Then

            InventoryToolStripMenuItem1.Visible = True
            ReportsToolStripMenuItem.Visible = True

            AdjustToolStripMenuItem.Visible = False
            RegisterToolStripMenuItem2.Visible = False
            BuildItemsToolStripMenuItem.Visible = False
            DateAnalysisToolStripMenuItem.Visible = False
            InventoryLedgerToolStripMenuItem.Visible = False
            ExportToolStripMenuItem.Visible = False
            ExpirationReportToolStripMenuItem.Visible = False
            Exit Sub








        Else
            ListToolStripMenuItem.Visible = False
            ReceiveToolStripMenuItem.Visible = False
            ReleaseToolStripMenuItem.Visible = False
            IntraTransferToolStripMenuItem.Visible = False
            InventoryToolStripMenuItem1.Visible = False
            ReportsToolStripMenuItem.Visible = False
            SetupToolStripMenuItem.Visible = False
            InventoryListToolStripMenuItem.Visible = False



        End If

    End Sub

    Sub closeThisForm()
        login_page.tbxUser.Text = ""
        With login_page
            .tbxUser.Text = ""
            .tbxPass.Text = ""

            login_page.Show()
            login_page.Focus()
            login_page.Select()
            Me.Dispose()
        End With

        closeforms()
    End Sub

    Sub closeforms()
        If MdiChildren.Length > 0 Then
            For Each r As Form In Me.MdiChildren
                r.Close()
            Next
        End If
    End Sub
    'Private Sub InventoryToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    closeforms()
    '    inventoryView.MdiParent = Me
    '    inventoryView.Show()
    'End Sub

    Private Sub Home_Page_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        If (MessageBox.Show("Are you sure you want to log out?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                e.Cancel = True
            Else
                e.Cancel = False
                login_page.tbxUser.Text = ""
                login_page.Show()
                login_page.Focus()
            End If

    End Sub

    Private Sub Home_Page_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            If MdiChildren.Length > 0 Then
                alreadyClose = False
            Else
                If (MessageBox.Show("Are you sure you want to log out?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
                Else
                    closeThisForm()
                    alreadyClose = True
                End If
            End If
        End If
    End Sub

    Private Sub ReceiveGoodsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReceiveGoodsToolStripMenuItem.Click
        closeforms()

        recvGoodsMain.MdiParent = Me
        recvGoodsMain.show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ReceiveRegisterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReceiveRegisterToolStripMenuItem.Click
        closeforms()
        recvListing.MdiParent = Me
        recvListing.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ReleaseGoodsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReleaseGoodsToolStripMenuItem.Click
        closeforms()
        releaseGoods.MdiParent = Me
        releaseGoods.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ReleaseRegisterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReleaseRegisterToolStripMenuItem.Click
        closeforms()
        releaseRegister.MdiParent = Me
        releaseRegister.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub AreaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AreaToolStripMenuItem.Click
        closeforms()
        area.MdiParent = Me
        area.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub DocumentsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DocumentsToolStripMenuItem.Click
        closeforms()
        documentTypes.MdiParent = Me
        documentTypes.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub RolesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RolesToolStripMenuItem.Click
        closeforms()
        roles.MdiParent = Me
        roles.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub LegerToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub BatchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BatchToolStripMenuItem.Click
        closeforms()
        viewBatch.MdiParent = Me
        viewBatch.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub LocationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LocationToolStripMenuItem.Click
        closeforms()
        viewLocations.MdiParent = Me
        viewLocations.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub UsersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UsersToolStripMenuItem.Click
        closeforms()
        users.MdiParent = Me
        users.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub TransferGoodsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TransferGoodsToolStripMenuItem1.Click
        closeforms()
        intraTransfer.MdiParent = Me
        intraTransfer.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub TransferRegisterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TransferRegisterToolStripMenuItem.Click
        closeforms()
        intraTransferRegister.MdiParent = Me
        intraTransferRegister.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub AddNewToolStripMenuItem2_Click(sender As Object, e As EventArgs)
        closeforms()
        inventoryAutoBuild.MdiParent = Me
        inventoryAutoBuild.Show()
    End Sub

    Private Sub RegisterToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        closeforms()
        inventoryGoods.MdiParent = Me
        inventoryGoods.Show()
    End Sub


    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        closeforms()
        inventoryAutoBuild.MdiParent = Me
        inventoryAutoBuild.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub RegisterToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles RegisterToolStripMenuItem1.Click
        closeforms()
        inventoryGoods.MdiParent = Me
        inventoryGoods.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub RegisterToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles RegisterToolStripMenuItem3.Click
        closeforms()
        adjustInventoryRegister.MdiParent = Me
        adjustInventoryRegister.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub NewToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem1.Click
        closeforms()
        adjustInventory.MdiParent = Me
        adjustInventory.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ItemRegisterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ItemRegisterToolStripMenuItem.Click
        closeforms()
        inventoryView.MdiParent = Me
        inventoryView.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub NewToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem2.Click
        closeforms()
        goodsAutoBuild.MdiParent = Me
        goodsAutoBuild.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ListToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ListToolStripMenuItem2.Click
        closeforms()
        goodsList.MdiParent = Me
        goodsList.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub RegisterToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles RegisterToolStripMenuItem4.Click
        closeforms()
        goodsRegister.MdiParent = Me
        goodsRegister.Show()
        MenuStrip1.Enabled = False
    End Sub


    Private Sub DateAnalysisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DateAnalysisToolStripMenuItem.Click
        closeforms()
        dateAnalysis.MdiParent = Me
        dateAnalysis.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub DocumentTypesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DocumentTypesToolStripMenuItem.Click
        closeforms()
        displayDocutypes.MdiParent = Me
        displayDocutypes.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub GeneralLedgerSummaryToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ViewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewToolStripMenuItem.Click
        closeforms()
        consolidatedInventoryList.MdiParent = Me
        consolidatedInventoryList.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub GeneralLedgerSpecificToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ItemListSummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) 
        closeforms()
        itemListSmmary.MdiParent = Me
        itemListSmmary.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ReserveGoodsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReserveGoodsToolStripMenuItem.Click
        closeforms()
        reserveRelease.MdiParent = Me
        reserveRelease.Show()
    End Sub

    'Private Sub tbxDisplayUser_TextChanged(sender As Object, e As EventArgs)
    '    tbxDisplayUser.Text = login_page.tbxUser.Text
    'End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs)
    End Sub

    Private Sub ConsolidatedInvLocationDetailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConsolidatedInvLocationDetailsToolStripMenuItem.Click
        closeforms()
        consolidatedInvLocationDetails.MdiParent = Me
        consolidatedInvLocationDetails.Show()

        MenuStrip1.Enabled = False
    End Sub

    Private Sub UpdateDescriptionToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'closeforms()
        'UpdateDes.MdiParent = Me

    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub ExportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportToolStripMenuItem.Click
        'closeforms()
        Export.ShowDialog()


    End Sub

    Private Sub ExpirationReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExpirationReportToolStripMenuItem.Click
        'closeforms()
        Expirationdatereport.ShowDialog()


    End Sub

    Private Sub UpdateExpirationDateToolStripMenuItem_Click(sender As Object, e As EventArgs)
        'closeforms()

    End Sub



    Public Sub MDIBGColor()
        Dim ctl As Control
        Dim ctlMDI As MdiClient

        ' Loop through all of the form's controls looking
        ' for the control of type MdiClient.
        For Each ctl In Me.Controls
            Try

                ' Attempt to cast the control to type MdiClient.
                ctlMDI = CType(ctl, MdiClient)

                ' Set the BackColor of the MdiClient control.
                ctlMDI.BackColor = Me.BackColor

            Catch exc As InvalidCastException
                ' Catch and ignore the error if casting failed.
            End Try
        Next

    End Sub

    Private Sub ReasonCodeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReasonCodeToolStripMenuItem.Click
        reasoncode.ShowDialog()
    End Sub
    Private Sub ReOrderToolStripMenuItem_Click(sender As Object, e As EventArgs)
        reorderingpoint.ShowDialog()
    End Sub


    Private Sub EditLocationToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles EditLocationToolStripMenuItem1.Click
        Loacationedit.ShowDialog()
    End Sub

    Private Sub UpdateDescriptionToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles UpdateDescriptionToolStripMenuItem1.Click
        UpdateDes.ShowDialog()
    End Sub

    Private Sub OrderPointToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OrderPointToolStripMenuItem.Click
        reordersetup.ShowDialog()
    End Sub

    Private Sub ReOrderToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReOrderToolStripMenuItem1.Click
        reorderingpoint.ShowDialog()
    End Sub

    Private Sub DetailedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DetailedToolStripMenuItem.Click
        closeforms()
        IGLR.MdiParent = Me
        IGLR.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub SummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SummaryToolStripMenuItem.Click
        closeforms()
        IGLRSummary.MdiParent = Me
        IGLRSummary.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub SpecificToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpecificToolStripMenuItem.Click
        closeforms()
        IGLRSpecific.MdiParent = Me
        IGLRSpecific.Show()
        MenuStrip1.Enabled = False
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        OPreport.ShowDialog()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        updatemasterlist.ShowDialog()
    End Sub


    Public Sub fetchnotif1()
        exdate.MdiParent = Me
        exdate.Show()
    End Sub

    Private Sub NotificationsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NotificationsToolStripMenuItem.Click
        exdate.MdiParent = Me
        exdate.Show()


    End Sub

    Private Sub ImportMasterListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportMasterListToolStripMenuItem.Click
        addnewitemlist.ShowDialog()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick

        'If roleId = "20020" Or "20016" Then
        '    Timer2.Enabled = False ' Disable the timer to prevent multiple form openings
        '    exdate.MdiParent = Me

        '    exdate.Show()

        'Else
        '    'nothing'
        'End If

    End Sub

    Private Sub ItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ItemToolStripMenuItem.Click
        StatusReport.ShowDialog()
    End Sub
End Class