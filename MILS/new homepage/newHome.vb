Imports System.Runtime.Remoting.Contexts
Imports System.Security
Imports System.Net

Public Class newHome
    Public userId As String
    Public roleId As String
    Public areaId As String
    Public empname As String
    Public branch As String
    Public notif1 As Boolean = False
    Private q As New qry
    Public alreadyClose As Boolean = False

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



    Private Sub NewHome_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        BarStaticItem2.Caption = "User: " & empname & "    " + "Area: " & branch
        BarStaticItem3.Caption = "Date: " + Date.Now.ToString("MM/dd/yyyy")
        Label1.Text = " IP Address: " & GetIPv4Address() + "    " + " Version 3.0"

        If roleId = 20020 Then

            Timer1.Enabled = True

        ElseIf roleId = 20016 Then

            Timer1.Enabled = True
        Else

            Timer1.Enabled = False

        End If

    End Sub

    Public Sub hidemodules()
        If roleId = 20020 Then
            AccordionControlElement1.Visible = True
            AccordionControlElement4.Visible = True
            AccordionControlElement5.Visible = True
            AccordionControlElement11.Visible = True
            AccordionControlElement15.Visible = True
            AccordionControlElement17.Visible = True
            AccordionControlElement18.Visible = True
            AccordionControlElement20.Visible = True
            AccordionControlElement56.Visible = True
            AccordionControlElement16.Visible = True
            AccordionControlElement59.Visible = True

        ElseIf roleId = 20016 Then
            AccordionControlElement1.Visible = True
            AccordionControlElement4.Visible = True
            AccordionControlElement5.Visible = True
            AccordionControlElement11.Visible = True
            AccordionControlElement15.Visible = True
            AccordionControlElement17.Visible = True
            AccordionControlElement18.Visible = True
            AccordionControlElement20.Visible = True
            AccordionControlElement56.Visible = False
            AccordionControlElement16.Visible = True

            AccordionControlElement2.Visible = False
            AccordionControlElement6.Visible = False
            AccordionControlElement8.Visible = False
            AccordionControlElement21.Visible = False
            AccordionControlElement23.Visible = False
            AccordionControlElement43.Visible = False
            AccordionControlElement45.Visible = False
            AccordionControlElement46.Visible = False
            AccordionControlElement47.Visible = False
            AccordionControlElement59.Visible = False
        ElseIf roleId = 20017 Then
            AccordionControlElement1.Visible = True
            AccordionControlElement4.Visible = True
            AccordionControlElement5.Visible = True
            AccordionControlElement11.Visible = True
            AccordionControlElement17.Visible = True
            AccordionControlElement56.Visible = False
            AccordionControlElement59.Visible = False
        ElseIf roleId = 20021 Then
            AccordionControlElement15.Visible = True
            AccordionControlElement31.Visible = True
            AccordionControlElement11.Visible = True
            AccordionControlElement14.Visible = True

            AccordionControlElement28.Visible = False
            AccordionControlElement29.Visible = False
            AccordionControlElement30.Visible = False
            AccordionControlElement32.Visible = False
            AccordionControlElement33.Visible = False
            AccordionControlElement10.Visible = False
            AccordionControlElement12.Visible = False
            AccordionControlElement13.Visible = False
            AccordionControlElement59.Visible = False
        ElseIf roleId = 20018 Then
            AccordionControlElement59.Visible = False
            AccordionControlElement56.Visible = False
            AccordionControlElement1.Visible = True
            AccordionControlElement4.Visible = True
            AccordionControlElement5.Visible = True
            AccordionControlElement11.Visible = True

        ElseIf roleId = 20019 Then
            AccordionControlElement17.Visible = True
            AccordionControlElement56.Visible = False
            AccordionControlElement59.Visible = False
        End If
    End Sub

    Private Sub AccordionControlElement2_Click(sender As Object, e As EventArgs) Handles AccordionControlElement2.Click
        recvGoodsMain.ShowDialog()
    End Sub

    Private Sub AccordionControlElement3_Click(sender As Object, e As EventArgs) Handles AccordionControlElement3.Click
        recvListing.ShowDialog()
    End Sub

    Private Sub AccordionControlElement6_Click(sender As Object, e As EventArgs) Handles AccordionControlElement6.Click
        releaseGoods.ShowDialog()
    End Sub

    Private Sub AccordionControlElement7_Click(sender As Object, e As EventArgs) Handles AccordionControlElement7.Click
        releaseRegister.ShowDialog()

    End Sub

    Private Sub AccordionControlElement8_Click(sender As Object, e As EventArgs) Handles AccordionControlElement8.Click
        intraTransfer.ShowDialog()
    End Sub

    Private Sub AccordionControlElement9_Click(sender As Object, e As EventArgs) Handles AccordionControlElement9.Click
        intraTransferRegister.ShowDialog()
    End Sub

    Private Sub AccordionControlElement21_Click(sender As Object, e As EventArgs) Handles AccordionControlElement21.Click
        adjustInventory.ShowDialog()
    End Sub

    Private Sub AccordionControlElement22_Click(sender As Object, e As EventArgs) Handles AccordionControlElement22.Click
        adjustInventoryRegister.ShowDialog()
    End Sub

    Private Sub AccordionControlElement23_Click(sender As Object, e As EventArgs) Handles AccordionControlElement23.Click
        inventoryAutoBuild.ShowDialog()
    End Sub

    Private Sub AccordionControlElement24_Click(sender As Object, e As EventArgs) Handles AccordionControlElement24.Click
        inventoryGoods.ShowDialog()
    End Sub

    Private Sub AccordionControlElement25_Click(sender As Object, e As EventArgs) Handles AccordionControlElement25.Click
        goodsAutoBuild.ShowDialog()
    End Sub

    Private Sub AccordionControlElement27_Click(sender As Object, e As EventArgs) Handles AccordionControlElement27.Click
        goodsList.ShowDialog()
    End Sub

    Private Sub AccordionControlElement26_Click(sender As Object, e As EventArgs) Handles AccordionControlElement26.Click
        goodsRegister.ShowDialog()
    End Sub

    Private Sub AccordionControlElement14_Click(sender As Object, e As EventArgs) Handles AccordionControlElement14.Click
        inventoryView.ShowDialog()
    End Sub

    Private Sub AccordionControlElement28_Click(sender As Object, e As EventArgs) Handles AccordionControlElement28.Click
        dateAnalysis.ShowDialog()
    End Sub

    Private Sub AccordionControlElement34_Click(sender As Object, e As EventArgs) Handles AccordionControlElement34.Click
        IGLR.ShowDialog()
    End Sub

    Private Sub AccordionControlElement38_Click(sender As Object, e As EventArgs) Handles AccordionControlElement38.Click
        IGLRSpecific.ShowDialog()
    End Sub

    Private Sub AccordionControlElement35_Click(sender As Object, e As EventArgs) Handles AccordionControlElement35.Click
        IGLRSummary.ShowDialog()
    End Sub

    Private Sub AccordionControlElement30_Click(sender As Object, e As EventArgs) Handles AccordionControlElement30.Click
        Expirationdatereport.ShowDialog()
    End Sub

    Private Sub AccordionControlElement37_Click(sender As Object, e As EventArgs) Handles AccordionControlElement37.Click
        OPreport.ShowDialog()
    End Sub

    Private Sub AccordionControlElement36_Click(sender As Object, e As EventArgs) Handles AccordionControlElement36.Click
        reorderingpoint.ShowDialog()
    End Sub

    Private Sub AccordionControlElement32_Click(sender As Object, e As EventArgs) Handles AccordionControlElement32.Click
        StatusReport.ShowDialog()
    End Sub

    Private Sub AccordionControlElement33_Click(sender As Object, e As EventArgs) Handles AccordionControlElement33.Click
        Export.ShowDialog()
    End Sub

    Private Sub AccordionControlElement39_Click(sender As Object, e As EventArgs) Handles AccordionControlElement39.Click
        consolidatedInventoryList.ShowDialog()
    End Sub

    Private Sub AccordionControlElement40_Click(sender As Object, e As EventArgs) Handles AccordionControlElement40.Click
        consolidatedInvLocationDetails.ShowDialog()
    End Sub

    Private Sub AccordionControlElement20_Click(sender As Object, e As EventArgs) Handles AccordionControlElement20.Click
        exdate.Show() ' Show the form.
        exdate.BringToFront() ' Bring the form to the front.
    End Sub

    Private Sub AccordionControlElement56_Click(sender As Object, e As EventArgs)
        Form1.ShowDialog()
    End Sub

    Private Sub AccordionControlElement53_Click(sender As Object, e As EventArgs) Handles AccordionControlElement53.Click
        viewBatch.ShowDialog()
    End Sub

    Private Sub AccordionControlElement54_Click(sender As Object, e As EventArgs) Handles AccordionControlElement54.Click
        displayDocutypes.ShowDialog()
    End Sub

    Private Sub AccordionControlElement55_Click(sender As Object, e As EventArgs) Handles AccordionControlElement55.Click
        viewLocations.ShowDialog()
    End Sub

    Private Sub AccordionControlElement41_Click(sender As Object, e As EventArgs) Handles AccordionControlElement41.Click
        area.ShowDialog()
    End Sub

    Private Sub AccordionControlElement42_Click(sender As Object, e As EventArgs) Handles AccordionControlElement42.Click
        documentTypes.ShowDialog()
    End Sub

    Private Sub AccordionControlElement43_Click(sender As Object, e As EventArgs) Handles AccordionControlElement43.Click
        roles.ShowDialog()
    End Sub

    Private Sub AccordionControlElement44_Click(sender As Object, e As EventArgs) Handles AccordionControlElement44.Click
        users.ShowDialog()
    End Sub

    Private Sub AccordionControlElement45_Click(sender As Object, e As EventArgs) Handles AccordionControlElement45.Click
        reasoncode.ShowDialog()
    End Sub

    Private Sub AccordionControlElement46_Click(sender As Object, e As EventArgs) Handles AccordionControlElement46.Click
        reordersetup.ShowDialog()
    End Sub

    Private Sub AccordionControlElement48_Click(sender As Object, e As EventArgs) Handles AccordionControlElement48.Click
        Loacationedit.ShowDialog()
    End Sub

    Private Sub AccordionControlElement49_Click(sender As Object, e As EventArgs) Handles AccordionControlElement49.Click
        UpdateDes.ShowDialog()
    End Sub

    Private Sub AccordionControlElement50_Click(sender As Object, e As EventArgs) Handles AccordionControlElement50.Click
        EditExpirationdate.ShowDialog()
    End Sub

    Private Sub AccordionControlElement51_Click(sender As Object, e As EventArgs) Handles AccordionControlElement51.Click
        updatemasterlist.ShowDialog()
    End Sub



    Private Sub AccordionControlElement18_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        If roleId = "20020" Or "20016" Then

            Timer1.Enabled = False ' Disable the timer to prevent multiple form openings

            exdate.Show() ' Show the form.
            exdate.BringToFront() ' Bring the form to the front.
        Else
            'nothing'
        End If

    End Sub

    Private Sub AccordionControlElement16_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AccordionControlElement57_Click(sender As Object, e As EventArgs) Handles AccordionControlElement57.Click
        TileSetup.ShowDialog()
    End Sub

    Private Sub NewHome_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If (MessageBox.Show("Are you sure you want to log out?", "Info", MessageBoxButtons.YesNo) = System.Windows.Forms.DialogResult.No) Then
            e.Cancel = True
        Else
            e.Cancel = False
            login_page.tbxUser.Text = ""
            login_page.Show()
            login_page.Focus()
        End If

        exdate.Close()
    End Sub

    Private Sub FluentDesignFormControl1_Click(sender As Object, e As EventArgs) Handles FluentDesignFormControl1.Click

    End Sub

    Private Sub SkinDropDownButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles SkinDropDownButtonItem1.ItemClick
    End Sub

    Private Sub AccordionControlElement58_Click(sender As Object, e As EventArgs) Handles AccordionControlElement58.Click
        Aging.Show()
    End Sub

    Private Sub FluentDesignFormContainer1_Click(sender As Object, e As EventArgs) Handles FluentDesignFormContainer1.Click
        exdate.Close()
    End Sub

    Private Sub AccordionControlElement56_Click_1(sender As Object, e As EventArgs) Handles AccordionControlElement56.Click
        Form1.ShowDialog()
    End Sub

    Private Sub AccordionControlElement59_Click(sender As Object, e As EventArgs) 

    End Sub

    Private Sub AccordionControlElement59_Click_1(sender As Object, e As EventArgs) Handles AccordionControlElement59.Click

    End Sub

    Private Sub AccordionControlElement64_Click(sender As Object, e As EventArgs) Handles AccordionControlElement64.Click
        Mrupload.ShowDialog()
    End Sub

    Private Sub AccordionControlElement65_Click(sender As Object, e As EventArgs)
        Mrtvupload.ShowDialog()
    End Sub

    Private Sub AccordionControlElement63_Click(sender As Object, e As EventArgs) Handles AccordionControlElement63.Click
        addnewitemlist.ShowDialog()
    End Sub

    Private Sub AccordionControlElement52_Click(sender As Object, e As EventArgs) Handles AccordionControlElement52.Click
        MainVendor.ShowDialog()
    End Sub
End Class
