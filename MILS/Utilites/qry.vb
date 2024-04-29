Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Drawing.Image
Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes
Imports System.Data.OleDb
Imports System
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class qry
    Public SQL As New sqlcon
    Private f As New functions

    Public path As String = "\\172.16.7.7\MILS_v3\PMNI\MILS-NPMI\MILS\"
    'Public path As String = "\\172.16.10.218\Users\HSDP_SYS_DEV\Desktop\WORK\Evermore\MILS-Evermore\MILS\"
    Public Sub reportlog()
        Dim RPuser1 As String = "HSDP_DEPOT"
        Dim RPpass2 As String = "123456$hsdp"

        With Export
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With Aging
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With Expirationdatereport
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With IGLR
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With IGLRSpecific
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With dateAnalysis
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With IGLRSummary
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With StatusReport
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

        With OPreport
            .RPuser3 = RPuser1.ToString
            .RPpass3 = RPpass2.ToString
        End With

    End Sub
    Public Sub logInUser(un As String, pass As String)
        'If un = "0000" And pass = "0000" Then
        '    With login_page
        '        .tbxUser.Text = ""
        '        .tbxPass.Text = ""
        '        .lblError.Visible = False
        '        .tbxUser.Select()
        '    End With
        '    Home_Page.Show() 'show home page form
        '    login_page.Hide() 'hide login form
        '    With Home_Page
        '        .userId = "0000"
        '        .roleId = "0000"
        '        .areaId = "0000"
        '    End With
        'Else
        SQL.AddParam("@un", un)
        SQL.AddParam("@pass", pass)

        SQL.ExecQueryDT("EXEC sp_loginUser @un = @un, @ps = @pass;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With newHome
                    .empname = ""
                    .userId = ""
                    .roleId = ""
                    .areaId = ""
                    .empname = r("empname")
                    .branch = r("branch")
                    .userId = r("empId")
                    .roleId = r("role_id")
                    .areaId = r("area_id")
                End With
            Next

            With login_page
                .tbxPass.Text = ""
            End With
            newHome.Show()
            newHome.hidemodules()
            login_page.Hide()
        Else
            MessageBox.Show("Invalid Account, try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            With login_page
                .tbxUser.Text = ""
                .tbxPass.Text = ""
                .tbxUser.Select()
            End With
        End If
    End Sub

    Public Sub addArea(an As String, ad As String, rd As Boolean)
        SQL.AddParam("@an", an)
        SQL.ExecQueryDT("SELECT * FROM tblAreas
                        WHERE area = @an;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            With areaAdd
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxAreaName.Text = ""
                .tbxAreaName.Select()
            End With
        Else
            SQL.AddParam("@an", an.ToUpper)
            SQL.AddParam("@ad", ad)
            SQL.AddParam("@rd", rd)
            SQL.AddParam("@encBy", newHome.userId)
            SQL.ExecQueryDT("INSERT INTO tblAreas
                            (area, [description], encBy,lastModified,isInternal)
                            VALUES
                            (@an,@ad,@encBy,GETDATE(),@rd);")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Area successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With areaAdd
                .tbxAreaName.Text = ""
                .tbxAreaDes.Text = ""
                .rdInternal.Checked = True
                .lblErr.Visible = False
                .tbxAreaName.Select()
            End With
        End If
    End Sub 'add Area

    Public Sub loadArea(lv As ListView, tbx As TextBox, cb As String)

        If cb = "External" Then
            cb = 0
        Else
            cb = 1
        End If
        SQL.AddParam("@flter", tbx.Text & "%")
        SQL.AddParam("@varea", cb)

        SQL.ExecQueryDT("
                        SELECT id,area,description,encBy,lastModified,case  
	                        when isInternal = '1' then CAST('Internal' as varchar(225))
	                        else CAST('External' as varchar(225))
	                        end as isinternal
	                        FROM tblAreas
	                        where  isInternal = @varea AND area LIKE @flter

")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(5))
                End With
            Next
        End If
    End Sub

    Public Sub loadallxArea(lv As ListView, tbx As TextBox)
        SQL.AddParam("@flter", tbx.Text & "%")

        SQL.ExecQueryDT("SELECT * FROM tblAreas WHERE area LIKE @flter;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    If r(5) = "1" Then
                        .SubItems.Add("Internal")
                    Else
                        .SubItems.Add("External")
                    End If
                End With
            Next
        End If
    End Sub 'load list view
    'Public Sub loadcbcat(cb As ComboBox)
    '    cb.Items.Clear()
    '    SQL.ExecQueryDT("select  isInternal from tblAreas")
    '    If SQL.HasException(True) Then Exit Sub

    '    If SQL.RecordCountDT > 0 Then
    '        For Each r As DataRow In SQL.DBDT.Rows
    '            If r(0) = "1" Then
    '                cb.Items.Add("Internal")
    '            Else
    '                cb.Items.Add("External")
    '            End If

    '        Next
    '    Else
    '        Exit Sub
    '    End If
    'End Sub 'load list view

    Public Sub fetchArea(id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT * FROM tblAreas
                        WHERE id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With areaDetails
                    .tbxAreaName.Text = r(1)
                    .tbxAreaDes.Text = r(2)
                    If r(5) = "1" Then
                        .rdInternal.Checked = True
                    ElseIf r(5) = "0" Then
                        .rdExternal.Checked = True
                    End If
                End With
            Next
        Else
        End If
    End Sub 'fetch Area to Update

    Public Sub editArea(id As String, an As String, ad As String, ii As Boolean)
        SQL.AddParam("@id", id)
        SQL.AddParam("@an", an)
        SQL.ExecQueryDT("SELECT * FROM tblAreas
                        WHERE id <> @id AND area = @an;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                MessageBox.Show("Data already exists.", "Invalid", MessageBoxButtons.OK, MessageBoxIcon.Error)
                With areaDetails
                    .tbxAreaName.Text = ""
                    .tbxAreaName.Select()
                End With
                Exit Sub
            Next
        Else
            SQL.AddParam("@id", id)
            SQL.AddParam("@an", an.ToUpper)
            SQL.AddParam("@ad", ad)
            SQL.AddParam("@ii", ii)
            SQL.AddParam("@encBy", newHome.userId)

            SQL.ExecQueryDT("UPDATE tblAreas
                            SET area = @an
	                            ,[description] = @ad
	                            ,encBy = @encBy
	                            ,lastModified = GETDATE()
	                            ,isInternal = @ii
                            WHERE id = @id;")
            If SQL.HasException(True) Then Exit Sub
        End If
        MessageBox.Show("Successfully updated.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        With areaDetails
            .closeForm()
        End With
        With area
            .loadForm()
            .Focus()
            .Select()
        End With

    End Sub 'Update area

    Public Sub suggestDocType(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT remarks FROM tblDocumentTypes WHERE isIn = 1 and isPrimary = 1;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub fetchIdDocType(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes WHERE isIn = 1 AND isPrimary = 1 AND remarks = @t;")
        If SQL.HasException(True) Then Exit Sub
        With recvGoodsMain
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = ""
                    .docTypeId = r(0)
                Next
            Else
                .docTypeId = ""
                .tbxDocType.Text = ""
            End If
        End With
        With recvListingDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = ""
                    .docTypeId = r(0)
                Next
            Else
                .docTypeId = ""
                .tbxDocType.Text = ""
            End If
        End With

    End Sub

    'Get The Reference Document Code
    Public Sub fetchreferencename(id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("Select docName from tblDocumentTypes where id = @id")
        If SQL.HasException(True) Then Exit Sub

        With recvListingDetails

            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docCode = r(0)
                Next
            End If

        End With

        With recvGoodsMain

            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docCode = r(0)
                Next
            End If

        End With

    End Sub

    Public Sub suggestDocTypeRecvListing(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT docName FROM tblDocumentTypes
                        WHERE isIn = 1 and isPrimary = 1;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub fetchIdDocTypeListing(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE isIn = 1 AND isPrimary = 1 AND docName = @t;")
        If SQL.HasException(True) Then Exit Sub
        With recvListing
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = r(0)
                Next

            End If
        End With
    End Sub

    Public Sub suggestDocTypeRelsListing(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT docName FROM tblDocumentTypes
                        WHERE isIn = 0 and isPrimary = 1;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub fetchIdDocTypeListingRels(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE isIn = 0 AND isPrimary = 1 AND docName = @t;")
        If SQL.HasException(True) Then Exit Sub
        With releaseRegister
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = r(0)
                Next
            Else
                .docTypeId = ""
                .tbxDocType.Text = ""
            End If
        End With
    End Sub

    Public Sub suggestRefDocType(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT remarks FROM tblDocumentTypes
                        WHERE isIn = 1 and isPrimary = 0;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub fetchIdRefDocType(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
                    WHERE isIn = 1 and isPrimary = 0 And remarks =  @t;")
        If SQL.HasException(True) Then Exit Sub
        With recvGoodsMain
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docRefTypeId = r(0)
                Next
            Else
                .docRefTypeId = ""
                .tbxRefDocType.Text = ""
            End If
        End With

        With recvListingDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docRefTypeId = r(0)
                Next
            Else
                .docRefTypeId = ""
                .tbxRefDocType.Text = ""
            End If
        End With
    End Sub

    'Select Sender in Receiving
    Public Sub loadSender(cbx As ComboBox)
        cbx.Items.Clear()

        SQL.AddParam("@aId", newHome.areaId)

        SQL.ExecQueryDS("SELECT area FROM tblAreas
                        WHERE id NOT IN (@aId)
                        ORDER BY area;")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDS > 0 Then

            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    'Select Vendor in Receiving
    Public Sub loadSenderVendor(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDS("Select Vendor from tblvendor")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    'Series for document number in Receiving for Recieving Report
    Public Sub loaddocnum(DocType As String)

        SQL.AddParam("@Doctype", DocType)

        SQL.ExecQueryDS("Select remarks from tblDocumentTypes where remarks =@Doctype")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDS > 0 Then



            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                With recvGoodsMain
                    .tbxDocNum.Text = r(0)
                End With
            Next

        Else

        End If
    End Sub
    Public Sub loadSenderRegister(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.AddParam("@aId", newHome.areaId)
        SQL.ExecQueryDS("SELECT area FROM tblAreas 
                        ORDER BY area;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    'fetch id of sender for Receiving Report
    Public Function fetchSenderId(ByVal t As String)
        Dim id As String = ""

        SQL.AddParam("@txt", t)
        SQL.ExecQueryDT("SELECT id FROM tblAreas
                        WHERE area = @txt;")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                id = r(0)
            Next
        End If

        Return id
    End Function

    'fetch id of vendor for Receiving Report
    Public Function fetchSenderIdVendor(ByVal t As String)
        Dim id As String = ""

        SQL.AddParam("@txt", t)
        SQL.ExecQueryDT("Select id from tblvendor WHERE Vendor = @txt;")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                id = r(0)
            Next
        End If

        Return id
    End Function
    'Public Sub loadLVforRecvGoods(lv As ListView, tbx As TextBox, des As TextBox, colSelected As Integer)
    '    If colSelected = 0 Then
    '        'SQL.AddParam("@flter", tbx.Text & "%")
    '        'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description
    '        'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id LIKE @flter AND goodDes LIKE @flter1 ORDER BY id ASC;")
    '        SQL.AddParam("@flter", tbx.Text)
    '        SQL.AddParam("@flter1", des.Text)
    '        SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")
    '        If SQL.HasException(True) Then Exit Sub
    '        lv.Items.Clear()
    '        lv.View = View.Details
    '        lv.FullRowSelect = True
    '        lv.GridLines = True
    '        Dim itemNew As New ListViewItem
    '        If SQL.RecordCountDT <> 0 Then
    '            For Each r As DataRow In SQL.DBDT.Rows
    '                itemNew = lv.Items.Add(r(0))
    '                With itemNew
    '                    .SubItems.Add(r(1))
    '                End With
    '            Next
    '        End If
    '    ElseIf colSelected = 1 Then
    '        'SQL.AddParam("@flter", "%" & tbx.Text & "%")
    '        'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description
    '        'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE goodDes LIKE @flter AND goodes LIKE @flter1 ORDER BY id ASC;")
    '        SQL.AddParam("@flter", tbx.Text)
    '        SQL.AddParam("@flter1", des.Text)
    '        SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")
    '        If SQL.HasException(True) Then Exit Sub
    '        lv.Items.Clear()
    '        lv.View = View.Details
    '        lv.FullRowSelect = True
    '        lv.GridLines = True
    '        Dim itemNew As New ListViewItem
    '        If SQL.RecordCountDT <> 0 Then
    '            For Each r As DataRow In SQL.DBDT.Rows
    '                itemNew = lv.Items.Add(r(0))
    '                With itemNew
    '                    .SubItems.Add(r(1))
    '                End With
    '            Next
    '        End If
    '    Else

    '    End If
    'End Sub


    Public Sub loadLVforRecvGoods(lv As ListView, tbx As TextBox, tbx2 As TextBox, colSelected As Integer)

        If colSelected = 0 Then
            SQL.AddParam("@fltr", tbx.Text)
            SQL.AddParam("@fltr2", "%" & tbx2.Text & "%")
            'SQL.AddParam("@flter", "%" & tbx.Text & "%")											
            'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description											

            'LAG IN EXECUTION OF QUERY TO DISPLAY ALL PRODUCTS RE: SELECT * FROM statement..											
            'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id LIKE @flter OR goodDes LIKE @flter1 ORDER BY id ASC;")											
            'SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")											
            'SQL.ExecQueryDT("Select * from tblGoods where id LIKE @flter")											
            SQL.ExecQueryDT("
                                SELECT TOP 50 '0' AS DUMMY ,T1.id,T1.goodDes 
                                FROM tblGoods T1
                                WHERE (T1.Status = 1) AND (T1.id LIKE '%'+ @fltr+'%' OR @fltr='') 
                                AND (T1.goodDes LIKE @fltr2 OR @fltr2 ='')
                                ORDER BY  T1.id ASC 

                            ")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True

            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                    End With
                Next
            End If

        ElseIf colSelected = 1 Then
            SQL.AddParam("@flter", tbx.Text)
            'SQL.AddParam("@flter1", des.Text)											
            'SQL.AddParam("@flter", "%" & tbx.Text & "%")											
            'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description											
            'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id LIKE @flter OR goodes LIKE @flter1 ORDER BY id ASC;")											
            'SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")											
            SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True

            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                    End With
                Next
            End If
        Else

        End If
    End Sub



    Public Sub loadLVforRecvGoodsDes(lv As ListView, des As TextBox, colSelected As Integer)

        If colSelected = 0 Then

            SQL.AddParam("@flter", "%" & des.Text & "%")
            'SQL.AddParam("@flter", "%" & tbx.Text & "%")										
            'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description										

            'LAG IN EXECUTION OF QUERY TO DISPLAY ALL PRODUCTS RE: SELECT * FROM statement..										
            'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id LIKE @flter OR goodDes LIKE @flter1 ORDER BY id ASC;")										
            'SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")										
            'SQL.ExecQueryDT("Select * from tblGoods where id LIKE @flter")										

            SQL.ExecQueryDT("EXEC sp_loadDescription @flter = @flter")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True

            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                    End With
                Next
            End If

        ElseIf colSelected = 1 Then
            SQL.AddParam("@flter", des.Text)
            'SQL.AddParam("@flter1", des.Text)										
            'SQL.AddParam("@flter", "%" & tbx.Text & "%")										
            'SQL.AddParam("@flter1", "%" & des.Text & "%") 'added tbx for description										
            'SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id LIKE @flter OR goodes LIKE @flter1 ORDER BY id ASC;")										
            'SQL.ExecQueryDT("EXEC sp_loadGoods @flter = @flter, @flter1 = @flter1;")										
            SQL.ExecQueryDT("EXEC sp_loadDescription @flter = @flter")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True

            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                    End With
                Next
            End If
        Else

        End If
    End Sub


    Public Sub loadLVforRecvBatches(lv As ListView, tbx As TextBox)
        SQL.AddParam("@flter", tbx.Text & "%")
        SQL.ExecQueryDT("SELECT * FROM tblbatches WHERE batchName LIKE @flter  ORDER BY batchName;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub


    Public Sub loadforinventorycountencoder(tbx As String)
        SQL.AddParam("@id", tbx)
        SQL.ExecQueryDT("select name + ' '+ lastname from tblEmpDetails where id = @id")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With adjustInventoryRegisterDetails
                For Each r As DataRow In SQL.DBDT.Rows
                    .TextBox1.Text = r(0)
                Next
            End With
        Else

        End If
    End Sub
    Public Sub addBatch(name As String, des As String)
        SQL.AddParam("@name", name)
        SQL.ExecQueryDT("SELECT * FROM tblbatches WHERE batchName = @name;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With addNewBatch
                .tbxBatchName.Text = ""
                .tbxBatchName.Select()
                .lblErr.Visible = True
                .lblErr.Text = "Data already exist."
            End With
            Exit Sub

        Else
            SQL.AddParam("@des", des)
            SQL.AddParam("@encBy", newHome.userId)
            SQL.AddParam("@name", name.ToUpper())

            SQL.ExecQueryDT("INSERT INTO tblBatches
                                    (batchName,remarks,encBy,encDate)
                                    VALUES
                                    (@name,@des,@encBy,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            addNewBatch.closeThisForm()
        End If
    End Sub

    Public Sub loadLVRecvLocations(lv As ListView, tbx As TextBox)
        SQL.AddParam("@flter", tbx.Text & "%")
        SQL.AddParam("@aId", newHome.areaId)
        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE [name] LIKE @flter AND areaId = @aId
                    ORDER BY [name];")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub

    Public Sub addLocations(name As String, rmks As String)
        SQL.AddParam("@name", name)
        SQL.AddParam("@aId", newHome.areaId)
        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE [name] = @name AND areaId = @aId;")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Location successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
        If SQL.RecordCountDT <> 0 Then
            With addNewLocation
                .tbxLocName.Text = ""
                .tbxLocName.Select()
                .lblErr.Visible = True
                .lblErr.Text = "Data already exist."
            End With
            Exit Sub
        Else
            SQL.AddParam("@name", name.ToUpper)
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@encBy", newHome.userId)
            SQL.AddParam("@aId", newHome.areaId)
            SQL.ExecQueryDT("INSERT INTO tblLocations
                            ([name], remarks,encby,areaId,lastModified)
                            VALUES
                            (@name,@rmks,@encBy,@aId,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            addNewLocation.closeThisForm()
        End If
    End Sub

    Public Sub validateDocumentBeforAddingRecvTrans(docType As String, docTypeId As String, docNum As String, area As String)

        SQL.AddParam("@docType", docTypeId)
        SQL.AddParam("@docNo", docNum)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblRecvHeaders T0
                        WHERE docType = @docType
                        AND docNo = @docNo
                        AND (SELECT area_id FROM tblEmpDetails WHERE id = T0.encBy) = @area;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            If (MessageBox.Show(docType & " #" & docNum & " has been previously used. " & vbCrLf & vbCrLf &
                                "Click OK if you wish to record the transaction using a duplicate " & docType & " #." & vbCrLf &
                                "Click CANCEL if you do not want to record the transaction.",
                                "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.Cancel) Then
                With recvGoodsMain
                    .tbxDocType.Text = ""
                    .tbxDocNum.Text = ""
                    .tbxDocType.Select()
                End With
                Exit Sub
            Else

            End If
        End If


    End Sub

    Public Sub addRecvTransaction(docType As String,
                                  docNo As String,
                                  docDate As Date,
                                  senderId As String,
                                  remarks As String,
                                  oType As String,
                                  docRefType As String,
                                  docRefNum As String,
                                  encBy As String,
                                  area As String, series As String, remarks2 As String)



        Dim transId As String = ""
        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double = 0 'Integer change to Double for decimal values
        Dim xdate As String = ""

        SQL.AddParam("@docType", docType)
        SQL.AddParam("@docNum", docNo)
        SQL.AddParam("@docDate", docDate)
        SQL.AddParam("@sId", senderId)
        SQL.AddParam("@rmks", remarks)
        SQL.AddParam("@oType", oType)
        SQL.AddParam("@docRefType", docRefType)
        SQL.AddParam("@docRefNum", docRefNum)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@areaId", area)
        SQL.AddParam("@series", series)
        SQL.AddParam("@remarks2", remarks2)

        SQL.ExecQueryDT("INSERT INTO tblRecvHeaders
                        (docType, docNo, docDate, encDate, encBy,sender,remarks,ownershipType,docRefType,docRefNo,areaId,series,remarks2)
                        VALUES
                        (@docType, @docNum, @docDate, GETDATE(),@encBy, @sId, @rmks, @oType, @docRefType, @docRefNum,@areaId,@series,@remarks2);
                        SELECT SCOPE_IDENTITY();")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transId = r(0)
                'MsgBox(transId)
            Next
        End If

        With recvGoodsMain.dgvRecv
            For i As Integer = 0 To .Rows.Count - 2 Step +1
                SQL.AddParam("@entryId", transId)
                SQL.AddParam("@gId", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@bId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@lId", .Rows(i).Cells(2).Value.ToString)
                SQL.AddParam("@aId", area)
                SQL.AddParam("@qty", .Rows(i).Cells(6).Value.ToString)
                SQL.AddParam("@exdate", .Rows(i).Cells(7).Value.ToString)

                SQL.ExecQueryDT("INSERT INTO tblRecvDetails
                                (transId, goodId,batchId,areaId,locId,qty,Expirationdate)
                                VALUES
                                (@entryId,@gId,@bId,@aId,@lId,@qty,@exdate);")
                If SQL.HasException(True) Then Exit Sub
            Next
        End With

        SQL.AddParam("@entryId", transId)
        SQL.ExecQueryDT("SELECT * FROM tblRecvDetails
                        WHERE transId = @entryId;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                goodId = r("goodId")
                batchId = r("batchId")
                areaId = r("areaId")
                locId = r("locId")
                qty = r("qty")
                xdate = r("Expirationdate")
                SQL.AddParam("@gId", goodId)
                SQL.AddParam("@bId", batchId)
                SQL.AddParam("@aId", areaId)
                SQL.AddParam("@lId", locId)
                SQL.AddParam("@qty", qty)
                SQL.AddParam("@exdate", xdate)
                SQL.ExecQueryDT("SELECT * FROM tblProducts
                                    WHERE goodId = @gId
                                    AND areaId = @aId
                                    AND batchId = @bId
                                    AND locationId = @lId
                                    AND Expirationdate = @exdate")
                If SQL.HasException(True) Then Exit Sub
                If SQL.RecordCountDT > 0 Then
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@aId", areaId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)





                    SQL.ExecQueryDT("UPDATE tblProducts
								set qty = qty + @qty
                                WHERE goodId = @gId
								AND areaId = @aId
								AND batchId = @bId
								AND locationId = @lId
                                AND Expirationdate = @exdate")

                    If SQL.HasException(True) Then Exit Sub
                Else
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@aId", areaId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)
                    SQL.ExecQueryDT("INSERT INTO tblProducts
								(goodId,areaId,batchId,locationId,qty,Expirationdate)
								VALUES
								(@gId,@aId,@bId,@lId,@qty,@exdate);")
                    If SQL.HasException(True) Then Exit Sub
                End If
            Next
        End If

        MessageBox.Show("Entry Successfully Added")
        With recvGoodsMain
            .dgvRecv.AllowUserToAddRows = False
            .dgvRecv.Enabled = False
            .tbxEntry.Text = transId
            .btnAdd.Text = "Add New Entry"
            .Button1.Visible = True
            .btnAdd.Select()
            .lblErr.Visible = True
            .lblErr.ForeColor = Color.Green

        End With

    End Sub

    Public Sub loadLVRecvListing(lv As ListView, sender As String, type As Boolean, dateFrom As Date, dateTo As Date, docType As String, docNum As String, area As String, roleId As String, areaName As String)
        'supervisor roleId 2016 change to roleId 20017 and manager roleId 20018 change to 20016 
        If roleId = "20016" Or roleId = "20020" Then
            SQL.AddParam("@dFrom", dateFrom)
            SQL.AddParam("@dTo", dateTo)
            SQL.AddParam("@sender", sender)
            SQL.AddParam("@docType", docType & "%")
            SQL.AddParam("@docNum", docNum & "%")
            SQL.AddParam("@area", areaName)
            SQL.ExecQueryDT("SELECT 
                                    '1' as dummy
                                    ,T0.docDate
                                    ,T0.id
                                    ,T2.docName
                                    ,T0.docNo
                                    ,T3.area
                                    ,T4.area
                                    ,CONVERT(DATE, T0.encDate)
                                    ,T1.name+' '+T1.middlename+' '+ T1.lastname
                                    ,t0.series
                                    FROM tblRecvHeaders T0
                                    INNER JOIN tblEmpDetails T1 ON T1.id=T0.encBy
                                    INNER JOIN tblDocumentTypes T2 ON T2.id=T0.docType
                                    INNER JOIN tblAreas T3 ON T3.id=T0.sender
                                    INNER JOIN tblAreas T4 ON T4.id=T0.areaId
                                    WHERE (T0.docDate BETWEEN @dFrom AND @dTo)
                                    AND (T0.docType LIKE @docType)
                                    AND (T0.docNo LIKE @docNum)
                                    AND (T3.area = @sender OR @sender='')
                                    AND (T4.area = @area OR @area ='')
                                    ORDER BY T0.encDate")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            Try
                If SQL.RecordCountDT <> 0 Then
                    For Each r As DataRow In SQL.DBDT.Rows
                        itemNew = lv.Items.Add(r(0))
                        With itemNew
                            If Not IsDBNull(r(1)) Then
                                .SubItems.Add(r(1))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(2)) Then
                                .SubItems.Add(r(2))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(3)) Then
                                .SubItems.Add(r(3))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(4)) Then
                                .SubItems.Add(r(4))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(5)) Then
                                .SubItems.Add(r(5))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(6)) Then
                                .SubItems.Add(r(6))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(7)) Then
                                .SubItems.Add(r(7))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(8)) Then
                                .SubItems.Add(r(8))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(9)) Then
                                .SubItems.Add(r(9))
                            Else
                                .SubItems.Add("")
                            End If
                        End With
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else

            SQL.AddParam("@dFrom", dateFrom)
            SQL.AddParam("@dTo", dateTo)
            SQL.AddParam("@sender", sender)
            SQL.AddParam("@docType", docType & "%")
            SQL.AddParam("@docNum", docNum & "%")
            SQL.AddParam("@area", area)
            SQL.ExecQueryDT("SELECT 
                                    '1' as dummy
                                    ,T0.docDate
                                    ,T0.id
                                    ,T2.docName
                                    ,T0.docNo
                                    ,T3.area
                                    ,T4.area
                                    ,CONVERT(DATE, T0.encDate)
                                    ,T1.name+' '+T1.middlename+' '+ T1.lastname
                                    ,t0.series
                                    FROM tblRecvHeaders T0
                                    INNER JOIN tblEmpDetails T1 ON T1.id=T0.encBy
                                    INNER JOIN tblDocumentTypes T2 ON T2.id=T0.docType
                                    INNER JOIN tblAreas T3 ON T3.id=T0.sender
                                    INNER JOIN tblAreas T4 ON T4.id=T0.areaId
                                    WHERE (T0.docDate BETWEEN @dFrom AND @dTo)
                                    AND (T0.docType LIKE @docType)
                                    AND (T0.docNo LIKE @docNum)
                                    AND (T3.area = @sender OR @sender='')
                                    AND (T4.id = @area)
                                    ORDER BY T0.encDate")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            Try
                If SQL.RecordCountDT <> 0 Then
                    For Each r As DataRow In SQL.DBDT.Rows
                        itemNew = lv.Items.Add(r(0))
                        With itemNew
                            If Not IsDBNull(r(1)) Then
                                .SubItems.Add(r(1))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(2)) Then
                                .SubItems.Add(r(2))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(3)) Then
                                .SubItems.Add(r(3))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(4)) Then
                                .SubItems.Add(r(4))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(5)) Then
                                .SubItems.Add(r(5))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(6)) Then
                                .SubItems.Add(r(6))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(7)) Then
                                .SubItems.Add(r(7))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(8)) Then
                                .SubItems.Add(r(8))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(9)) Then
                                .SubItems.Add(r(9))
                            Else
                                .SubItems.Add("")
                            End If
                        End With
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Public Sub loadLVRecvDetails(lv As ListView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT '' as dummy,T0.goodId,T1.goodDes,T2.batchName,T3.name,T0.qty,T0.Expirationdate
                            FROM tblRecvDetails T0 
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                            INNER JOIN tblLocations T3 ON T0.locId=T3.id
                            WHERE T0.transId = @id")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    If r.IsNull(6) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(6))
                    End If
                End With
            Next
        End If
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT
                            T0.id
                            ,(SELECT area FROM tblAreas WHERE id = T0.sender)
                            ,(SELECT remarks FROM tblDocumentTypes WHERE id = T0.docType)
                            ,T0.docNo
                            ,T0.docDate
                            ,(SELECT remarks FROM tblDocumentTypes WHERE id = T0.docRefType)
                            ,T0.docRefNo
                            ,T0.ownershipType
                            ,T0.remarks
                            ,(SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  FROM tblEmpDetails WHERE id = T0.encBy)
                        FROM tblRecvHeaders T0
                        WHERE id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With recvListingDetails
                    If Not IsDBNull(r(0)) Then
                        .tbxEntry.Text = r(0)
                    Else
                        .tbxEntry.Text = ""
                    End If

                    If Not IsDBNull(r(1)) Then
                        .cbxSender.Text = r(1)
                    Else
                        .cbxSender.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDocType.Text = r(2)
                    Else
                        .tbxDocType.Text = ""
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxDocNum.Text = r(3)
                    Else
                        .tbxDocNum.Text = ""
                    End If

                    If Not IsDBNull(r(4)) Then
                        .dtpRefDate.Value = r(4)
                    Else
                        .dtpRefDate.Value = Date.Now
                    End If

                    If Not IsDBNull(r(5)) Then
                        .tbxRefDocType.Text = r(5)
                    Else
                        .tbxRefDocType.Text = ""
                    End If

                    If Not IsDBNull(r(6)) Then
                        .tbxRefDocNum.Text = r(6)
                    Else
                        .tbxRefDocNum.Text = ""
                    End If

                    If Not IsDBNull(r(7)) Then
                        .cbxOwnership.Text = r(7)
                    Else
                        .cbxOwnership.Text = ""
                    End If

                    If Not IsDBNull(r(8)) Then
                        .tbxRemarks.Text = r(8)
                    Else
                        .tbxRemarks.Text = ""
                    End If

                    If Not IsDBNull(r(9)) Then
                        .tbxEmployeeName.Text = r(9)
                    Else
                        .tbxEmployeeName.Text = ""
                    End If
                End With
            Next
        End If
    End Sub


    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub loadLVRecvDetail(dg As DataGridView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT t0.id,T0.goodId,T1.goodDes,T2.batchName,T3.name,T0.qty, COALESCE(T0.Expirationdate, 'N/A') AS Expirationdate,T0.locId,T0.batchId
                            FROM tblRecvDetails T0 
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                            INNER JOIN tblLocations T3 ON T0.locId=T3.id
                            WHERE T0.transId = @id")
        If SQL.HasException(True) Then Exit Sub
        With recvListingDetails.dg1
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT
            .ClearSelection()
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .Columns(0).Visible = False
            .Columns(1).HeaderText = "PRODUCT NUMBER"
            .Columns(2).HeaderText = "DESCRIPTION"
            .Columns(3).HeaderText = "BATCH"
            .Columns(4).HeaderText = "LOCATION"
            .Columns(5).HeaderText = "QUANTITY"
            .Columns(6).HeaderText = "EXPIRATION DATE"
            .Columns(7).Visible = False
            .Columns(8).Visible = False

            .Columns(1).Width = 150
            .Columns(2).Width = 530
            .Columns(3).Width = 120
            .Columns(4).Width = 120
            .Columns(5).Width = 100
            .Columns(6).Width = 170


            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = False
            .Columns(6).ReadOnly = True

            '.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9, FontStyle.Bold)
                .Columns(Index).DefaultCellStyle.Font = New Font(.DefaultCellStyle.Font.FontFamily, 9)
                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .CurrentCell = Nothing
        End With




        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT
                            T0.id
                            ,(SELECT area FROM tblAreas WHERE id = T0.sender)
                            ,(SELECT remarks FROM tblDocumentTypes WHERE id = T0.docType)
                            ,T0.docNo
                            ,T0.docDate
                            ,(SELECT remarks FROM tblDocumentTypes WHERE id = T0.docRefType)
                            ,T0.docRefNo
                            ,T0.ownershipType
                            ,T0.remarks
                            ,(SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,'')) FROM tblEmpDetails WHERE id = T0.encBy)
                            ,T0.remarks2
                        FROM tblRecvHeaders T0
                        WHERE id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With recvListingDetails
                    If Not IsDBNull(r(0)) Then
                        .tbxEntry.Text = r(0)
                    Else
                        .tbxEntry.Text = ""
                    End If

                    If Not IsDBNull(r(1)) Then
                        .cbxSender.Text = r(1)
                    Else
                        .cbxSender.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDocType.Text = r(2)
                    Else
                        .tbxDocType.Text = ""
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxDocNum.Text = r(3)
                    Else
                        .tbxDocNum.Text = ""
                    End If

                    If Not IsDBNull(r(4)) Then
                        .dtpRefDate.Value = r(4)
                    Else
                        .dtpRefDate.Value = Date.Now
                    End If

                    If Not IsDBNull(r(5)) Then
                        .tbxRefDocType.Text = r(5)
                    Else
                        .tbxRefDocType.Text = ""
                    End If

                    If Not IsDBNull(r(6)) Then
                        .tbxRefDocNum.Text = r(6)
                    Else
                        .tbxRefDocNum.Text = ""
                    End If

                    If Not IsDBNull(r(7)) Then
                        .cbxOwnership.Text = r(7)
                    Else
                        .cbxOwnership.Text = ""
                    End If

                    If Not IsDBNull(r(8)) Then
                        .tbxRemarks.Text = r(8)
                    Else
                        .tbxRemarks.Text = ""
                    End If

                    If Not IsDBNull(r(9)) Then
                        .tbxEmployeeName.Text = r(9)
                    Else
                        .tbxEmployeeName.Text = ""
                    End If

                    If Not IsDBNull(r(10)) Then
                        .TextBox1.Text = r(10)
                    Else
                        .TextBox1.Text = ""
                    End If

                End With
            Next
        End If
    End Sub
    Public Sub loadLVInventory(lv As ListView, area As String, id As String, des As String, com As String)

        If newHome.roleId = 20020 Or newHome.roleId = 20021 Then

            SQL.AddParam("@aId", area)
            SQL.AddParam("@id", id)
            SQL.AddParam("@des", "%" + des + "%")
            SQL.AddParam("@stat", com)
            SQL.ExecQueryDT("SELECT TOP 50 '' as dummy,T0.goodId,T1.goodDes,SUM(qty) as QTY,
                            CASE
                            WHEN T1.Status = 1 THEN 'Active'
                            WHEN T1.Status = 2 THEN 'Inactive'
                            WHEN T1.Status = 3 THEN 'Discontinued'
                            END AS StatusDisplay
                            FROM tblProducts T0
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            WHERE (T0.goodId like '%'+ @id +'%' or @id = '')
                            AND (t1.goodDes LIKE @des or @des ='')
                            AND (T1.Status = @stat or @stat ='')
                            GROUP BY T0.goodId,T1.goodDes,T1.Status
                           
                            ORDER BY t0.goodId ASC")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(4))
                    End With
                Next
            End If

        Else

            SQL.AddParam("@aId", area)
            SQL.AddParam("@id", id)
            SQL.AddParam("@des", "%" + des + "%")
            SQL.ExecQueryDT("SELECT TOP 50 T0.areaId,T0.goodId,T1.goodDes,SUM(qty) 
                            FROM tblProducts T0
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            WHERE (T0.goodId like '%'+ @id +'%'  or @id = '')
                            AND (t1.goodDes LIKE @des or @des ='')
                            AND (T0.areaId =@aId)
                            GROUP BY T0.areaId,T0.goodId,T1.goodDes
                      
                            ORDER BY t0.goodId ASC")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        '.SubItems.Add(r(4))
                        '.SubItems.Add(r(5))
                    End With
                Next
            End If
        End If

    End Sub

    Public Sub loadLVProductDetails(lv As ListView, area As String, id As String, batch As String, location As String)


        If newHome.roleId = 20020 Or newHome.roleId = 20021 Then
            SQL.AddParam("@id", id)
            SQL.AddParam("@batch", batch + "%")
            SQL.AddParam("@loc", location + "%")
            SQL.ExecQueryDT("SELECT T0.goodId,T1.goodDes,T2.batchName,T3.name,T0.qty FROM tblProducts T0
	                        INNER JOIN tblGoods T1 ON T0.goodId=T1.id
	                        INNER JOIN tblBatches T2 ON T0.batchId=T2.id
	                        INNER JOIN tblLocations T3 ON T0.locationId=T3.id
	                        INNER JOIN tblAreas T4 ON T0.areaId=T4.id
	                        WHERE (T2.batchName LIKE '%'+ @batch +'%' or @batch='')
	                        AND (T3.name LIKE '%'+ @loc +'%' or @loc='')
	                        AND (t0.goodId = @id   );")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(4))

                    End With
                Next
            End If

            SQL.AddParam("@aId", area)
            SQL.AddParam("@id", id)
            SQL.AddParam("@b", batch + "%")
            SQL.AddParam("@l", location + "%")
            SQL.ExecQueryDT("SELECT
                        SUM(qty)
                        FROM tblProducts
                        WHERE goodId like '%'+ @id +'%' 
                        AND areaId  = @aId
                        AND (SELECT batchName FROM tblBatches WHERE id = batchId)  LIKE @b
                        AND (SELECT [name] FROM tblLocations WHERE id = locationId) LIKE @l;")
            If SQL.HasException(True) Then Exit Sub

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    With invViewDetails
                        If IsDBNull(r(0)) Then
                            .tbxTotalQty.Text = "0"
                        Else
                            .tbxTotalQty.Text = r(0)
                        End If
                    End With
                Next
            Else
                Exit Sub
            End If
        Else
            SQL.AddParam("@aId", area)
            SQL.AddParam("@id", id)
            SQL.AddParam("@b", batch + "%")
            SQL.AddParam("@l", location + "%")
            SQL.ExecQueryDT("SELECT
	                        goodId
	                        ,(SELECT goodDes FROM tblGoods WHERE id = goodId)
	                        ,(SELECT batchName FROM tblBatches WHERE id = batchId) 
	                        ,(SELECT [name] FROM tblLocations WHERE id = locationId)
	                        ,T1.qty
	                        ,T1.qty - 
                        ISNULL((SELECT 1 * SUM(qty) FROM tblReserveReleaseDetails x1 
                        JOIN tblReserveReleaseHeaders x2 ON x2.id = x1.transId
                        WHERE prodId = T1.id AND x2.[status] = 'P'),0) AS [Available]
                        FROM tblProducts T1
                        WHERE goodId like '%'+ @aId +'%' 
                        AND areaId  = @aId
                        AND T1.qty > 0
                        AND (SELECT batchName FROM tblBatches WHERE id = batchId)  LIKE @b
                        AND (SELECT [name] FROM tblLocations WHERE id = locationId) LIKE @l
                        ORDER BY (SELECT [name] FROM tblLocations WHERE id = locationId);")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(4))
                        .SubItems.Add(r(5))
                    End With
                Next
            End If

            SQL.AddParam("@aId", area)
            SQL.AddParam("@id", id)
            SQL.AddParam("@b", batch + "%")
            SQL.AddParam("@l", location + "%")
            SQL.ExecQueryDT("SELECT
                        SUM(qty)
                        FROM tblProducts
                        WHERE goodId like '%'+ @id +'%' 
                        AND areaId  = @aId
                        AND (SELECT batchName FROM tblBatches WHERE id = batchId)  LIKE @b
                        AND (SELECT [name] FROM tblLocations WHERE id = locationId) LIKE @l;")
            If SQL.HasException(True) Then Exit Sub

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    With invViewDetails
                        If IsDBNull(r(0)) Then
                            .tbxTotalQty.Text = "0"
                        Else
                            .tbxTotalQty.Text = r(0)
                        End If
                    End With
                Next
            Else
                Exit Sub
            End If
        End If









    End Sub

    Public Sub loadLVDocTypes(lv As ListView, tbx1 As TextBox, tbx2 As TextBox)
        SQL.AddParam("@flter1", tbx1.Text & "%")
        SQL.AddParam("@flter2", tbx2.Text & "%")
        SQL.ExecQueryDT("SELECT * FROM tblDocumentTypes
                        WHERE (docName LIKE @flter1 AND remarks LIKE @flter2)
                        ORDER BY docName;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))

                    If r(5) = "1" Then
                        .SubItems.Add("Receiving")
                    Else
                        .SubItems.Add("Releasing")
                    End If

                    If r(6) = "1" Then
                        .SubItems.Add("Primary")
                    Else
                        .SubItems.Add("Secondary")
                    End If
                End With
            Next
        End If
    End Sub

    Public Sub addNewDocType(name As String, rmks As String, recv As Boolean, prim As Boolean)
        SQL.AddParam("@name", name)
        SQL.ExecQueryDT("SELECT * FROM tblDocumentTypes WHERE docName = @name;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With docuAdd
                .lblErr.Visible = True
                .lblErr.Text = "Data already exist."
                .tbxDocName.Text = ""
                .tbxDocName.Select()
            End With
        Else
            SQL.AddParam("@name", name.ToUpper())
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@isRecv", recv)
            SQL.AddParam("@isPrim", prim)
            SQL.AddParam("@encBy", newHome.userId)

            SQL.ExecQueryDT("INSERT INTO tblDocumentTypes
                            (docName,remarks,encby,[date],isIn,isPrimary)
                            VALUES
                            (@name,@rmks,@encBy,GETDATE(),@isRecv,@isPrim);")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Document type successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With docuAdd
                .closethisform()
            End With
        End If
    End Sub

    Public Sub fetchDocType(id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT * FROM tblDocumentTypes WHERE id = @id;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With docuDetails
                    If Not IsDBNull(r(1)) Then
                        .tbxDocName.Text = r(1)
                    Else
                        .tbxDocName.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDocRem.Text = r(2)
                    Else
                        .tbxDocRem.Text = ""
                    End If

                    If Not IsDBNull(r(5)) Then
                        If r(5) = "1" Then
                            .rdRecv.Checked = True
                        Else
                            .rdRels.Checked = True
                        End If
                    Else

                    End If

                    If Not IsDBNull(r(6)) Then
                        If r(6) = "1" Then
                            .rdPrim.Checked = True
                        Else
                            .rdSec.Checked = True
                        End If
                    Else

                    End If

                End With
            Next
        Else
            With docuDetails
                .lblErr.Visible = True
                .lblErr.Text = "No data found..."
            End With
        End If

    End Sub

    Public Sub UpdateDocType(id As String, name As String, rmks As String, recv As Boolean, prim As Boolean)
        SQL.AddParam("@id", id)
        SQL.AddParam("@name", name)

        SQL.ExecQueryDT("SELECT * FROM tblDocumentTypes WHERE id <> @id AND docName = @name;")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Updated successfully.", "Info:", MessageBoxButtons.OK, MessageBoxIcon.Information)
        If SQL.RecordCountDT <> 0 Then
            With docuDetails
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxDocName.Text = ""
                .tbxDocName.Select()
            End With
            Exit Sub
        Else
            With docuDetails
                .lblErr.Visible = False

                SQL.AddParam("@id", id)
                SQL.AddParam("@name", name)
                SQL.AddParam("@rmks", rmks)
                SQL.AddParam("@recv", recv)
                SQL.AddParam("@prim", prim)
                SQL.AddParam("@encBy", newHome.userId)

                SQL.ExecQueryDT("UPDATE tblDocumentTypes
                                SET docName = @name
                                ,remarks = @rmks
                                ,encby = @encBy
                                ,[date]= GETDATE()
                                ,isIn = @recv
                                ,isPrimary = @prim
                                WHERE
                                id =@id;")
                .closeThisForm()
            End With
        End If
    End Sub

    Public Sub loadLVRoles(lv As ListView, tbx As TextBox)
        SQL.AddParam("@filter", tbx.Text & "%")
        SQL.ExecQueryDT("SELECT * FROM tblRoles
                        WHERE role LIKE @filter
                        ORDER BY id")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))

                    If r.IsNull(2) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(2))
                    End If
                End With
            Next
        End If
    End Sub


    Public Sub addRoles(name As String, rmks As String)
        SQL.AddParam("@name", name)
        SQL.ExecQueryDT("SELECT * FROM tblRoles WHERE [role] = @name;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With rolesAdd
                .tbxRole.Text = ""
                .tbxRole.Select()
                .lblErr.Visible = True
                .lblErr.Text = "Data already exist."
            End With
        Else
            SQL.AddParam("@name", name.ToUpper)
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@encBy", newHome.userId)
            SQL.ExecQueryDT("INSERT INTO tblRoles
                            ([role],[description],encBy,lastModified)
                            VALUES
                            (@name,@rmks,@encBy,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Role successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With rolesAdd
                .closeThisForm()
            End With
        End If
    End Sub

    Public Sub fetchRoles(id As String)
        SQL.AddParam("@id", id)

        SQL.ExecQueryDT("SELECT * FROM tblRoles WHERE id = @id;")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With rolesDetails
                    If Not IsDBNull(r(1)) Then
                        .tbxRole.Text = ""
                        .tbxRole.Text = r(1)
                    Else
                        .tbxRole.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxRoleDes.Text = ""
                        .tbxRoleDes.Text = r(2)
                    Else
                        .tbxRoleDes.Text = ""
                    End If
                End With
            Next
        Else

        End If

    End Sub

    Public Sub updateRoles(id As String, name As String, rmks As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@name", name)
        SQL.ExecQueryDT("SELECT * FROM tblRoles WHERE id <> @id AND [role] = @name;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With rolesDetails
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxRole.Text = ""
                .tbxRole.Select()
            End With
        Else
            SQL.AddParam("@name", name.ToUpper)
            SQL.AddParam("@id", id)
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@encBy", newHome.userId)

            SQL.ExecQueryDT("UPDATE tblRoles
                            SET [role]= @name
                            ,[description]= @rmks
                            ,encBy = @encBy
                            ,lastModified = GETDATE()
                            WHERE
                            id= @id;")
            MessageBox.Show("Updated Successfully", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With rolesDetails
                .lblErr.Visible = False
                .closeThisForm()
            End With
        End If

    End Sub

    Public Sub suggestDocTypeRels(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT remarks FROM tblDocumentTypes
                        WHERE isIn = 0 and isPrimary = 1;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub suggestRefDocTypeRels(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT remarks FROM tblDocumentTypes
                        WHERE isIn = 0 and isPrimary = 0;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub fetchIdDocTypeRels(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE remarks = @t;")
        If SQL.HasException(True) Then Exit Sub
        With releaseGoods
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = ""
                    .docTypeId = r(0)
                Next
            Else
                .docTypeId = ""
                .tbxDocType.Text = ""
            End If
        End With
    End Sub

    Public Sub fetchIdDocTypeReservRels(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE remarks = @t;")
        If SQL.HasException(True) Then Exit Sub
        With reserveRelease
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = ""
                    .docTypeId = r(0)
                Next
            Else
                .docTypeId = ""
                .tbxDocType.Text = ""
            End If
        End With
    End Sub

    Public Sub fetchIdRefDocTypeRels(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE remarks = @t;")
        If SQL.HasException(True) Then Exit Sub
        With releaseGoods
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docRefTypeId = ""
                    .docRefTypeId = r(0)
                Next
            Else
                .docRefTypeId = ""
                .tbxRefDocType.Text = ""
            End If
        End With
    End Sub

    Public Sub fetchIdRefDocTypeReserveRels(t As String)
        SQL.AddParam("@t", t)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes
            WHERE remarks = @t;")
        If SQL.HasException(True) Then Exit Sub
        With reserveRelease
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docRefTypeId = ""
                    .docRefTypeId = r(0)
                Next
            Else
                .docRefTypeId = ""
                .tbxRefDocType.Text = ""
            End If
        End With
    End Sub

    Public Sub loadLVProducts(lv As ListView, txt As TextBox, des As TextBox)
        If newHome.roleId = 20020 Then
            SQL.AddParam("@fltr", txt.Text)
            SQL.AddParam("@fltr2", "%" & des.Text & "%")
            SQL.AddParam("@area", newHome.areaId)
            SQL.ExecQueryDT(" SELECT TOP 50 T0.id,T0.goodId,T0.areaId,T0.batchId,T0.locationId,T0.qty,T1.goodDes,T2.batchName,T3.name,T0.Expirationdate
                        FROM tblProducts T0
                        INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                        INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                        INNER JOIN tblLocations T3 ON T0.locationId=T3.id
                        WHERE T0.qty>0 
                        AND (T0.goodId like '%'+ @fltr +'%'   OR @fltr='')
                       AND (T1.goodDes like @fltr2 OR @fltr2='')                                      
                      ORDER BY T0.goodId ASC")

            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    'itemNew = lv.Items.Add(r("id")) 'id
                    'With itemNew
                    '    .SubItems.Add(r("goodId")) 'pdno
                    '    .SubItems.Add(r("goodDes")) 'description
                    '    .SubItems.Add(r("Batch")) 'batch
                    '    .SubItems.Add(r("Locations")) 'loc
                    '    .SubItems.Add(r("qty")) 'quantity
                    '    .SubItems.Add(r("batchId")) 'batch Id
                    '    .SubItems.Add(r("locationId")) 'locID
                    'End With
                    itemNew = lv.Items.Add(r(0)) 'id
                    With itemNew
                        .SubItems.Add(r(1)) 'pdno
                        .SubItems.Add(r(6)) 'description
                        .SubItems.Add(r(7)) 'batch
                        .SubItems.Add(r(8)) 'loc
                        .SubItems.Add(r(5)) 'quantity
                        .SubItems.Add(r(3)) 'batch Id
                        .SubItems.Add(r(4)) 'locID

                        If r.IsNull(9) Then
                            .SubItems.Add("N/A")
                        Else
                            .SubItems.Add(r(9))
                        End If
                    End With
                Next
            End If
        Else
            SQL.AddParam("@fltr", txt.Text)
            SQL.AddParam("@fltr2", "%" & des.Text & "%")
            SQL.AddParam("@area", newHome.areaId)
            SQL.ExecQueryDT("
                        SELECT TOP 50 T0.id,T0.goodId,T0.areaId,T0.batchId,T0.locationId,T0.qty,T1.goodDes,T2.batchName,T3.name,T0.Expirationdate
                        FROM tblProducts T0
                        INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                        INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                        INNER JOIN tblLocations T3 ON T0.locationId=T3.id
                        WHERE T0.qty>0 
                        AND (T0.goodId like '%'+ @fltr +'%' OR @fltr='')
                       AND (T1.goodDes like @fltr2 OR @fltr2='')                                      
                        AND (T0.areaId =@area) ORDER BY T0.goodId ASC")

            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    'itemNew = lv.Items.Add(r("id")) 'id
                    'With itemNew
                    '    .SubItems.Add(r("goodId")) 'pdno
                    '    .SubItems.Add(r("goodDes")) 'description
                    '    .SubItems.Add(r("Batch")) 'batch
                    '    .SubItems.Add(r("Locations")) 'loc
                    '    .SubItems.Add(r("qty")) 'quantity
                    '    .SubItems.Add(r("batchId")) 'batch Id
                    '    .SubItems.Add(r("locationId")) 'locID
                    'End With
                    itemNew = lv.Items.Add(r(0)) 'id
                    With itemNew
                        .SubItems.Add(r(1)) 'pdno
                        .SubItems.Add(r(6)) 'description
                        .SubItems.Add(r(7)) 'batch
                        .SubItems.Add(r(8)) 'loc
                        .SubItems.Add(r(5)) 'quantity
                        .SubItems.Add(r(3)) 'batch Id
                        .SubItems.Add(r(4)) 'locID

                        If r.IsNull(9) Then
                            .SubItems.Add("N/A")
                        Else
                            .SubItems.Add(r(9))
                        End If
                    End With
                Next
            End If
        End If

    End Sub




    Public Function validationQuantityProd(id As String, qty As Double)
        Dim bool As Boolean = True
        Dim qtyDB As Double = 0.0 'Integer change to Double for decimal values

        If String.IsNullOrWhiteSpace(id) Then
            MessageBox.Show("Please select prodcut first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        Else

            SQL.AddParam("@id", id)

            SQL.ExecQueryDT("SELECT qty FROM tblProducts
                            WHERE id =  @id")
            If SQL.HasException(True) Then Return Nothing

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    qtyDB = (r(0))
                Next
            End If

            Dim dgvQty As Double = qty
            Dim qtyInDB As Double = qtyDB 'Integer change to Double for decimal values
            If dgvQty > qtyInDB Then
                bool = False
            End If
        End If
        Return bool
    End Function

    Public Function validationQuantityProdedit(id As String, qty As Double, reldetail As String)
        Dim bool As Boolean = True
        Dim qtyDB As Double = 0.0 'Integer change to Double for decimal values
        Dim relid As Double = 0.0

        If String.IsNullOrWhiteSpace(id) Then
            MessageBox.Show("Please select prodcut first.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        ElseIf id <> reldetail Then

            SQL.AddParam("@id", id)
            SQL.ExecQueryDT("SELECT qty FROM tblProducts
                            WHERE id =  @id")
            If SQL.HasException(True) Then Return Nothing

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    qtyDB = (r(0))
                Next
            End If
            Dim dgvQty As Double = qty
            Dim qtyInDB As Double = qtyDB 'Integer change to Double for decimal values
            If dgvQty > qtyInDB Then
                bool = False
            End If

        Else

            SQL.AddParam("@id", reldetail)
            SQL.ExecQueryDT("SELECT qty FROM tblRelsDetails WHERE id =@id")
            If SQL.HasException(True) Then Return Nothing

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    relid = (r(0))
                Next
            End If
            SQL.AddParam("@id", id)
            SQL.ExecQueryDT("SELECT qty FROM tblProducts
                            WHERE id =  @id")
            If SQL.HasException(True) Then Return Nothing

            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    qtyDB = (r(0))
                Next
            End If


            Dim dgvQty As Double = qty
            Dim qtyInDB As Double = qtyDB 'Integer change to Double for decimal values

            Dim tot As Double = qtyDB + relid
            If dgvQty > tot Then
                bool = False
            End If

        End If
            Return bool
    End Function

    Public Sub validateDocumentRelsTransaction(docType As String, doctypeid As String, docNo As String, area As String)
        SQL.AddParam("@docType", doctypeid)
        SQL.AddParam("@docNo", docNo)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblRelsHeader T0
                        WHERE docType = @docType
                        AND docNo = @docNo
                        AND (SELECT area_id FROM tblEmpDetails WHERE id = T0.encBy) = @area;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            If (MessageBox.Show(docType & " #" & docNo & " has been previously used. " & vbCrLf & vbCrLf &
                                "Click OK if you wish to record the transaction using a duplicate " & docType & " #." & vbCrLf &
                                "Click CANCEL if you do not want to record the transaction.",
                                "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.Cancel) Then
                With releaseGoods
                    .tbxDocType.Text = ""
                    .tbxDocNum.Text = ""
                    .tbxDocType.Select()
                End With
                Exit Sub
            Else

            End If
        End If
    End Sub

    Public Sub validateDocumentReservRelsTransaction(docType As String, doctypeid As String, docNo As String, area As String)
        SQL.AddParam("@docType", doctypeid)
        SQL.AddParam("@docNo", docNo)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblRelsHeader T0
                        WHERE docType = @docType
                        AND docNo = @docNo
                        AND (SELECT area_id FROM tblEmpDetails WHERE id = T0.encBy) = @area;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            If (MessageBox.Show(docType & " #" & docNo & " has been previously used. " & vbCrLf & vbCrLf &
                                "Click OK if you wish to record the transaction using a duplicate " & docType & " #." & vbCrLf &
                                "Click CANCEL if you do not want to record the transaction.",
                                "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.Cancel) Then
                With reserveRelease
                    .tbxDocType.Text = ""
                    .tbxDocNum.Text = ""
                    .tbxDocType.Select()
                End With
                Exit Sub
            Else

            End If
        End If
    End Sub

    Public Sub addRelsTransaction(docType As String, docNo As String, docDate As Date, encBy As String, receiver As String,
                   rmrks As String, oType As String, docRefType As String, docRefNo As String, area As String, series1 As String)
        Dim transId As String = ""
        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double = 0 'String change to Double for decimal values

        SQL.AddParam("@docType", docType)
        SQL.AddParam("@docNo", docNo)
        SQL.AddParam("@docDate", docDate)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@receiver", receiver)
        SQL.AddParam("@rmrks", rmrks)
        SQL.AddParam("@oType", oType)
        SQL.AddParam("@docRefType", docRefType)
        SQL.AddParam("@docRefNo", docRefNo)
        SQL.AddParam("@areaId", area)
        SQL.AddParam("@series", series1)

        SQL.ExecQueryDT("INSERT INTO tblRelsHeader
                        (docType, docNo, docDate, encDate, encBy,receiver,remarks,ownershipType,docRefType,docRefNo,areaId,series)
                        VALUES
                        (@docType,@docNo ,@docDate ,GETDATE(), @encBy, @receiver, @rmrks, @oType, @docRefType, @docRefNo,@areaId,@series);
                        SELECT SCOPE_IDENTITY();")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transId = r(0)
                'MsgBox(transId)
            Next
        End If

        With releaseGoods.dgvRels
            For i As Integer = 0 To .Rows.Count - 2 Step +1
                SQL.AddParam("@entryId", transId)
                SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@gId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@bId", .Rows(i).Cells(6).Value.ToString)
                SQL.AddParam("@lId", .Rows(i).Cells(7).Value.ToString)
                SQL.AddParam("@aId", area)
                SQL.AddParam("@qty", .Rows(i).Cells(5).Value.ToString)
                SQL.AddParam("@exdate", .Rows(i).Cells(8).Value.ToString)
                SQL.ExecQueryDT("INSERT INTO tblRelsDetails
                                (transId,pId, goodId,batchId,areaId,locId,qty,Expirationdate)
                                VALUES
                                (@entryId,@pId,@gId,@bId,@aId,@lId,@qty,@exdate)")
                If SQL.HasException(True) Then Exit Sub
            Next
        End With

        SQL.AddParam("@entryId", transId)
        SQL.ExecQueryDT("SELECT * FROM tblRelsDetails
        WHERE transId = @entryId;")
        If SQL.HasException(True) Then Exit Sub
        Dim pId As String
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                pId = r("pId")
                qty = r("qty")

                SQL.AddParam("@pId", pId)
                SQL.AddParam("@qty", qty)

                SQL.ExecQueryDT("UPDATE tblProducts
                                SET qty = qty - @qty
                                WHERE id = @pId;")
                If SQL.HasException(True) Then Exit Sub
            Next
        End If
        MsgBox("Successfully Added")
        With releaseGoods
            .dgvRels.AllowUserToAddRows = False
            .dgvRels.Enabled = False
            .tbxEntry.Text = transId
            .Button1.Visible = True
            .btnAdd.Text = "Add New Entry"
            .btnAdd.Select()
        End With

    End Sub

    Public Sub addReserveRelsTransaction(docType As String, docNo As String, docDate As Date, encBy As String, receiver As String,
                   rmrks As String, oType As String, docRefType As String, docRefNo As String, area As String)
        Dim transId As String = ""
        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double = 0 'String change to Double for decimal values

        SQL.AddParam("@docType", docType)
        SQL.AddParam("@docNo", docNo)
        SQL.AddParam("@docDate", docDate)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@receiver", receiver)
        SQL.AddParam("@rmrks", rmrks)
        SQL.AddParam("@oType", oType)
        SQL.AddParam("@docRefType", docRefType)
        SQL.AddParam("@docRefNo", docRefNo)
        SQL.AddParam("@areaId", area)
        SQL.AddParam("@stats", "P")

        SQL.ExecQueryDT("INSERT INTO tblReserveReleaseHeaders
                        (docType, docNo, docDate, encDate, encBy,recvr,remarks,ownershipType,docRefType,docRefNo,areaId,[status])
                        VALUES
                        (@docType,@docNo ,@docDate ,GETDATE(), @encBy, @receiver, @rmrks, @oType, @docRefType, @docRefNo,@areaId,@stats);
                        SELECT SCOPE_IDENTITY();")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transId = r(0)
                'MsgBox(transId)
            Next
        End If

        With reserveRelease.dgvRels
            For i As Integer = 0 To .Rows.Count - 2 Step +1
                SQL.AddParam("@trnsId", transId)
                SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@gId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@bId", .Rows(i).Cells(6).Value.ToString)
                SQL.AddParam("@lId", .Rows(i).Cells(7).Value.ToString)
                SQL.AddParam("@aId", area)
                SQL.AddParam("@qty", .Rows(i).Cells(5).Value.ToString)

                SQL.ExecQueryDT("INSERT INTO tblReserveReleaseDetails
                                (transId,prodId,goodId,batchId,locid,areaId,qty)
                                VALUES
                                (@trnsId,@pId, @gId, @bId, @lId, @aId, @qty);")
                If SQL.HasException(True) Then Exit Sub
            Next
        End With

        'SQL.AddParam("@entryId", transId)
        'SQL.ExecQueryDT("SELECT * FROM tblReserveReleaseDetails
        'WHERE transId = @entryId;")
        'If SQL.HasException(True) Then Exit Sub
        'Dim pId As String
        'If SQL.RecordCountDT <> 0 Then
        '    For Each r As DataRow In SQL.DBDT.Rows
        '        pId = r("pId")
        '        qty = r("qty")

        '        SQL.AddParam("@pId", pId)
        '        SQL.AddParam("@qty", qty)

        '        SQL.ExecQueryDT("UPDATE tblProducts
        '                        SET qty = qty - @qty
        '                        WHERE id = @pId;")
        '        If SQL.HasException(True) Then Exit Sub
        '    Next
        'End If

        With reserveRelease
            .dgvRels.AllowUserToAddRows = False
            .dgvRels.Enabled = False
            .tbxEntry.Text = transId
            .btnAdd.Text = "Add New Entry"
            .btnAdd.Select()
            .lblErr.Visible = True
            .lblErr.ForeColor = Color.Green
            .lblErr.Text = "Entry Successfully Added."
        End With

    End Sub

    Public Sub loadLVRelsListing(lv As ListView, recvr As String, dFrom As Date, dTo As Date, type As Boolean, docType As String, docNum As String, area As String, roleId As String, areaName As String)
        'manager roleId 20018 change to 20016
        'If roleId = "20018" Or roleId = 20018 Thens
        Dim areaid = newHome.areaId


        If roleId = "20016" Or roleId = "20020" Then

            SQL.AddParam("@dFrom", dFrom)
            SQL.AddParam("@dTo", dTo)
            SQL.AddParam("@docType", "%" & docType)
            SQL.AddParam("@docNum", docNum)
            SQL.AddParam("@area", area)
            SQL.AddParam("@areaName", areaName)
            SQL.AddParam("@rcvr", recvr)
            SQL.ExecQueryDT("SELECT 
                                '1' as dummy
                                ,T0.docDate
                                ,T0.id
                                ,t2.docName
                                ,t0.docNo
                                ,t3.area
                                ,t4.area
                                ,CONVERT(DATE, T0.encDate)
                                ,T1.name +' '+ T1.middlename + ' '+ T1.lastname
                                ,series
                                FROM tblRelsHeader T0
                                INNER JOIN tblEmpDetails T1 ON T1.id=T0.encBy
                                INNER JOIN tblDocumentTypes T2 ON T2.id=T0.docType
                                INNER JOIN tblAreas T3 ON T3.id=T0.receiver
                                INNER JOIN tblAreas T4 ON T4.id=T0.areaId
                                WHERE (T0.docDate BETWEEN @dFrom AND @dTo )
                                AND (T0.docType LIKE '%' + @docType OR @docType='')
                                AND (T4.area = @areaName OR @areaName ='')
                                AND (T3.area = @rcvr OR @rcvr='')
                                AND (T0.docNo LIKE '%' + @docNum  OR @docNum='')
                                ")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            Try
                If SQL.RecordCountDT <> 0 Then
                    For Each r As DataRow In SQL.DBDT.Rows
                        itemNew = lv.Items.Add(r(0))
                        With itemNew
                            If Not IsDBNull(r(1)) Then
                                .SubItems.Add(r(1))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(2)) Then
                                .SubItems.Add(r(2))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(3)) Then
                                .SubItems.Add(r(3))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(4)) Then
                                .SubItems.Add(r(4))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(5)) Then
                                .SubItems.Add(r(5))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(6)) Then
                                .SubItems.Add(r(6))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(7)) Then
                                .SubItems.Add(r(7))
                            Else
                                .SubItems.Add("")
                            End If
                            If Not IsDBNull(r(8)) Then
                                .SubItems.Add(r(8))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(9)) Then
                                .SubItems.Add(r(9))
                            Else
                                .SubItems.Add("")
                            End If
                        End With
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

        Else


            SQL.AddParam("@dFrom", dFrom)
            SQL.AddParam("@dTo", dTo)
            SQL.AddParam("@docType", "%" & docType)
            SQL.AddParam("@docNum", "%" & docNum)
            SQL.AddParam("@area", area)
            SQL.AddParam("@areaid", areaid)
            SQL.AddParam("@rcvr", recvr)
            SQL.ExecQueryDT("SELECT 
                                '1' as dummy
                                ,T0.docDate
                                ,T0.id
                                ,t2.docName
                                ,t0.docNo
                                ,t3.area
                                ,t4.area
                                ,CONVERT(DATE, T0.encDate)
                                ,T1.name +' '+ T1.middlename + ' '+ T1.lastname
                                ,series
                                FROM tblRelsHeader T0
                                INNER JOIN tblEmpDetails T1 ON T1.id=T0.encBy
                                INNER JOIN tblDocumentTypes T2 ON T2.id=T0.docType
                                INNER JOIN tblAreas T3 ON T3.id=T0.receiver
                                INNER JOIN tblAreas T4 ON T4.id=T0.areaId
                                WHERE (T0.docDate BETWEEN @dFrom AND @dTo )
                                AND (T0.docType LIKE @docType OR @docType='')
                                AND (T4.id = @areaid)
                                AND (T3.area = @rcvr or @rcvr='')
                                AND (T0.docNo LIKE @docNum OR @docNum='')
                                ")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            Try
                If SQL.RecordCountDT <> 0 Then
                    For Each r As DataRow In SQL.DBDT.Rows
                        itemNew = lv.Items.Add(r(0))
                        With itemNew
                            If Not IsDBNull(r(1)) Then
                                .SubItems.Add(r(1))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(2)) Then
                                .SubItems.Add(r(2))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(3)) Then
                                .SubItems.Add(r(3))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(4)) Then
                                .SubItems.Add(r(4))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(5)) Then
                                .SubItems.Add(r(5))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(6)) Then
                                .SubItems.Add(r(6))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(7)) Then
                                .SubItems.Add(r(7))
                            Else
                                .SubItems.Add("")
                            End If
                            If Not IsDBNull(r(8)) Then
                                .SubItems.Add(r(8))
                            Else
                                .SubItems.Add("")
                            End If

                            If Not IsDBNull(r(9)) Then
                                .SubItems.Add(r(9))
                            Else
                                .SubItems.Add("")
                            End If
                        End With
                    Next
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


        End If

    End Sub

    Public Sub loadLVRelsDetails(id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT T0.id,T0.goodId,T1.goodDes,T2.batchName,T3.name,T0.qty,COALESCE(T0.Expirationdate, 'N/A') AS Expirationdate,T0.pId,T2.id,t3.id
                        FROM tblRelsDetails T0
                        INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                        INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                        INNER JOIN tblLocations T3 ON T0.locId=T3.id
                        WHERE transId = @id")
        If SQL.HasException(True) Then Exit Sub
        With releaseDetails.dg1
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT
            .ClearSelection()
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .Columns(0).Visible = False

            .Columns(1).HeaderText = "PRODUCT NUMBER"
            .Columns(2).HeaderText = "DESCRIPTION"
            .Columns(3).HeaderText = "BATCH"
            .Columns(4).HeaderText = "LOCATION"
            .Columns(5).HeaderText = "QUANTITY"
            .Columns(6).HeaderText = "EXPIRATION DATE"

            .Columns(7).Visible = False
            .Columns(8).Visible = False
            .Columns(9).Visible = False


            .Columns(1).Width = 150
            .Columns(2).Width = 530
            .Columns(3).Width = 120
            .Columns(4).Width = 120
            .Columns(5).Width = 100
            .Columns(6).Width = 170
            '.Columns(7).Width = 110


            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = False
            .Columns(6).ReadOnly = True
            '.Columns(7).ReadOnly = True

            '.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9, FontStyle.Bold)
                .Columns(Index).DefaultCellStyle.Font = New Font(.DefaultCellStyle.Font.FontFamily, 9)

                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .CurrentCell = Nothing
        End With





        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT T0.id,T1.area,T2.remarks,T0.docNo,T0.docDate,T3.remarks,T0.docRefNo,T0.ownershipType,T0.remarks,T4.name+' '+T4.middlename+' '+T4.lastname AS ENCODER,T0.areaId
                            FROM tblRelsHeader T0
                            INNER JOIN tblAreas T1 ON T0.receiver=T1.id
                            INNER JOIN tblDocumentTypes T2 ON T0.docType=T2.id
                            INNER JOIN tblDocumentTypes T3 ON T0.docRefType=T3.id
                            INNER JOIN tblEmpDetails T4 ON T0.encBy=T4.id
                            WHERE T0.id=@id")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With releaseDetails
                    If Not IsDBNull(r(0)) Then
                        .tbxEntry.Text = r(0)
                    Else
                        .tbxEntry.Text = "N/A"
                    End If

                    If Not IsDBNull(r(1)) Then
                        .cbxReceiver.Text = r(1)
                    Else
                        .cbxReceiver.Text = "N/A"
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDocType.Text = r(2)
                    Else
                        .tbxDocType.Text = "N/A"
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxDocNum.Text = r(3)
                    Else
                        .tbxDocNum.Text = "N/A"
                    End If

                    If Not IsDBNull(r(4)) Then
                        .dtpRefDate.Value = r(4)
                    Else
                        .dtpRefDate.Value = Date.Now
                    End If

                    If Not IsDBNull(r(5)) Then
                        .tbxRefDocType.Text = r(5)
                    Else
                        .tbxRefDocType.Text = "N/A"
                    End If

                    If Not IsDBNull(r(6)) Then
                        .tbxRefDocNum.Text = r(6)
                    Else
                        .tbxRefDocNum.Text = "N/A"
                    End If

                    If Not IsDBNull(r(7)) Then
                        .cbxOwnership.Text = r(7)
                    Else
                        .cbxOwnership.Text = "N/A"
                    End If

                    If Not IsDBNull(r(8)) Then
                        .tbxRemarks.Text = r(8)
                    Else
                        .tbxRemarks.Text = "N/A"
                    End If

                    If Not IsDBNull(r(9)) Then
                        .tbxEmployeeName.Text = r(9)
                    Else
                        .tbxEmployeeName.Text = "N/A"
                    End If

                    If Not IsDBNull(r(10)) Then
                        .area = r(10)
                    End If

                End With
            Next
        End If
    End Sub

    Public Sub loadLVBatchesView(lv As ListView, tbx As TextBox)
        SQL.AddParam("@flter", tbx.Text & "%")
        SQL.ExecQueryDT("SELECT * FROM tblbatches WHERE batchName LIKE @flter  ORDER BY batchName asc")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(4))
                End With
            Next
        End If
    End Sub

    Public Sub loadLVLocationsView(lv As ListView, tbx As TextBox, aId As String, rId As String)
        'supervisor roleId 2016 change to roleId 20017 
        'and manager roleId 20018 change to 20016 
        'and encoder roleId 20019 change to 20018 
        'If rId = "20018" Or rId = "20016" Or rId = "20019" Then
        If rId = "20016" Or rId = "20017" Or rId = "20018" Then
            SQL.AddParam("@flter", tbx.Text & "%")
            SQL.AddParam("@aId", aId)
            SQL.ExecQueryDT("SELECT *
                                FROM tblLocations
                                WHERE [name] LIKE @flter AND areaid = @aId 
                                ORDER BY
                                  CASE
                                    WHEN ISNUMERIC(name) = 1 THEN REPLICATE('0', 100 - LEN(name)) + name
                                    ELSE name
                                  END ASC;
                                ")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(5))
                    End With
                Next
            End If
        Else
            SQL.AddParam("@flter", tbx.Text & "%")
            SQL.AddParam("@aId", aId)
            SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE [name] LIKE @flter
                        AND areaId =  @aId;")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(5))
                    End With
                Next
            End If
        End If
    End Sub
    Public Sub loadLVLocationsViewadmin(lv As ListView, tbx As String, aId As String)
        SQL.AddParam("@flter", tbx)
        SQL.AddParam("@aId", aId)
        SQL.ExecQueryDT("SELECT TOP 50 t0.id,t0.name,t0.remarks,t0.lastModified,t1.area
                            FROM tblLocations t0
                            inner join tblAreas t1 on t0.areaId=t1.id
                         WHERE (t0.name LIKE '%' + @flter + '%' or @flter ='') AND (t1.area = @aId or @aId ='')
                            ORDER BY
                                CASE
                                WHEN ISNUMERIC(name) = 1 THEN REPLICATE('0', 100 - LEN(name)) + name
                                ELSE name
                                END ASC")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                End With
            Next
        End If
    End Sub




    Public Sub loadLVUsers(lv As ListView, tbx1 As TextBox, tbx2 As TextBox, cbx As String, aid As String)
        SQL.AddParam("@flter1", tbx1.Text & "%")
        SQL.AddParam("@flter2", tbx2.Text & "%")
        SQL.AddParam("@flter3", cbx & "%")
        'SQL.AddParam("@aId", aid)
        SQL.ExecQueryDT("	select t0.id as ID ,CONCAT(t0.name,' ',t0.middlename,' ',t0.lastname) as USERS
						,t1.description as ROLE
						,t2.area as AREA
						,CASE 
							WHEN t0.status = '1' then 'Active'
							WHEN t0.status = '0' then 'Inactive'
						 END as STATUS
						from tblEmpDetails t0 
						inner join tblRoles t1 on t0.role_id=t1.id
						inner join tblAreas t2 on t0.area_id=t2.id
                        WHERE t0.[name] LIKE @flter1
                        AND t0.lastname LIKE @flter2
                        AND t1.description LIKE @flter3
                        AND (t1.description != 'SPU')")
        'AND T0.area_id = @aId
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    If Not IsDBNull(r(2)) Then
                        .SubItems.Add(r(2))
                    Else
                        .SubItems.Add("")
                    End If

                    If Not IsDBNull(r(3)) Then
                        .SubItems.Add(r(3))
                    Else
                        .SubItems.Add("")
                    End If

                    .SubItems.Add(r(4))
                End With
            Next
        End If
    End Sub
    Public Sub loadroles(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select description from tblRoles where id != 20020")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub loadareainarea(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select DISTINCT 
case  
	when isInternal = '1' then CAST('Internal' as varchar(225))
	else CAST('External' as varchar(225))
	end as AREA
from tblAreas ")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub
        End If
    End Sub
    Public Sub loadAreaforUsers(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDS("SELECT area FROM tblAreas WHERE isInternal <> 0;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    Public Sub loadRolesforUsers(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDS("SELECT [role] FROM tblRoles where id != 20020;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    Public Function fetchAreaforUsers(ByVal txt As String)
        Dim id As String = ""

        SQL.AddParam("@txt", txt)
        SQL.ExecQueryDT("SELECT id FROM tblAreas
                        WHERE area = @txt;")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                id = r(0)
            Next
        End If

        Return id
    End Function

    Public Function fetchAreaforhome(ByVal txt As String)
        Dim id As String = ""

        SQL.AddParam("@id", txt)
        SQL.ExecQueryDT("SELECT area FROM tblAreas
                        WHERE id = @txt;")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                id = r(0)
            Next
        End If

        Return id
    End Function

    Public Function fetchRolesforUsers(txt As String)
        Dim id As String = ""

        SQL.AddParam("@txt", txt)
        SQL.ExecQueryDT("SELECT id FROM tblRoles WHERE [role] = @txt;")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                id = r(0)
            Next
        End If

        Return id
    End Function

    Public Sub addUser(name As String, mName As String, lName As String,
                       rId As String, aId As String, encBy As String, uName As String, pass As String)
        SQL.AddParam("@uName", uName)
        SQL.ExecQueryDT("SELECT * FROM tblUsers WHERE username = @uName;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With usersAdd
                .tbxUserName.Text = ""
                .tbxUserName.Select()
                MessageBox.Show("Data Already Exists...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End With
        Else
            SQL.AddParam("@name", name)
            SQL.AddParam("@mName", mName)
            SQL.AddParam("@lName", lName)
            SQL.AddParam("@rId", rId)
            SQL.AddParam("@aId", aId)
            SQL.AddParam("@encBy", encBy)
            SQL.AddParam("@uName", uName)
            SQL.AddParam("@pass", pass)
            SQL.ExecQueryDT("EXEC spAddEmpDetails
	                    @name = @name
	                    ,@mName = @mName
	                    ,@lName = @lName
	                    ,@stats = 1
	                    ,@rId = @rId
	                    ,@aId = @aId
	                    ,@encBy = @encBy
	                    ,@uName =@uname
	                    ,@pass = @pass;")
            If SQL.HasException(True) Then Exit Sub
            If MessageBox.Show("New user successfully added.", "Info:", MessageBoxButtons.OK, MessageBoxIcon.Information) Then
                '(MessageBox.Show("You want to add more user?", "Info", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = System.Windows.Forms.DialogResult.No) Then
                With usersAdd
                    .closeThisForm()
                End With
                'Else
                '    With usersAdd
                '        .clearFields()
                '        .TabControl1.SelectedIndex = 0
                '        .tbxFN.Select()
                '    End With
                '    MessageBox.Show("You want to add more user?", "Info:", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Public Sub fetchUser(id As String)
        SQL.AddParam("@id", id)

        SQL.ExecQueryDT("SELECT T0.id, T0.[name], T0.middlename, T0.lastname,
	                        (SELECT [role] FROM tblRoles WHERE id = T0.role_id) AS [ROLE],
	                        (SELECT area FROM tblAreas WHERE id = T0.area_id ) AS [AREA],
	                        T0.[status], T1.username, T1.pass
                        FROM tblEmpDetails T0
                        LEFT JOIN tblUsers T1 ON T1.empId = T0.id
                        WHERE T0.id = @id;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With usersDetails

                    Try
                        If Not IsDBNull(r(1)) Then
                            .tbxFN.Text = r(1)
                        Else
                            .tbxFN.Text = ""
                        End If

                        If Not IsDBNull(r(2)) Then
                            .tbxMN.Text = r(2)
                        Else
                            .tbxMN.Text = ""
                        End If

                        If Not IsDBNull(r(3)) Then
                            .tbxLN.Text = r(3)
                        Else
                            .tbxLN.Text = ""
                        End If

                        If Not IsDBNull(r(4)) Then
                            .cbxRole.Text = r(4)
                        Else
                            .cbxRole.SelectedIndex = -1
                        End If

                        If Not IsDBNull(r(5)) Then
                            .cbxArea.Text = r(5)
                        Else
                            .cbxArea.SelectedIndex = -1
                        End If

                        If r(6) = "1" Then
                            .rdActive.Checked = True
                        Else
                            .rdInactive.Checked = True
                        End If

                        If Not IsDBNull(r(7)) Then
                            .tbxUserName.Text = r(7)
                        Else
                            .tbxUserName.Text = ""
                        End If

                        If Not IsDBNull(r(8)) Then
                            .tbxPass.Text = r(8)
                        Else
                            .tbxPass.Text = ""
                        End If
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                End With
            Next
        End If
    End Sub

    Public Sub updateUsers(id As String, name As String, mName As String, lName As String, roleID As String, areID As String, stats As String, Uname As String, pas As String, encBy As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@uName", Uname)

        SQL.ExecQueryDT("SELECT *
                        FROM tblEmpDetails T0
                        LEFT JOIN tblUsers T1 ON T1.empId = T0.id
                        WHERE T0.id <> @id AND T1.username = @uName;")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Successfully updated...", "Info:", MessageBoxButtons.OK, MessageBoxIcon.Information)
        If SQL.RecordCountDT > 0 Then
            With usersDetails
                MessageBox.Show("Data Already Exists...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                .tbxUserName.Text = ""
                .tbxUserName.Select()
            End With
            Exit Sub
        Else
            SQL.AddParam("@name", name)
            SQL.AddParam("@mName", mName)
            SQL.AddParam("@lName", lName)
            SQL.AddParam("@stats", stats)
            SQL.AddParam("@rId", roleID)
            SQL.AddParam("@aId", areID)
            SQL.AddParam("@encBy", encBy)
            SQL.AddParam("@uName", Uname)
            SQL.AddParam("@pass", pas)
            SQL.AddParam("@id", id)

            SQL.ExecQueryDT("EXEC spUpdateEmpDetails
	                        @name = @name
	                        ,@mName = @mName
	                        ,@lName = @lName
	                        ,@stats = @stats
	                        ,@rId = @rId
	                        ,@aId = @aId
	                        ,@encBy = @encBy
	                        ,@uName = @uName
	                        ,@pass = @pass
	                        ,@id = @id;")
            If SQL.HasException(True) Then Exit Sub
            With usersDetails
                .closeThisForm()
            End With
        End If
    End Sub

    Public Sub fetchUserLogIn(id)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT CONCAT(CASE WHEN name IS NULL THEN '' ELSE name +' ' END,
	                        CASE WHEN middlename IS NULL THEN '' ELSE middlename +' ' END,
	                        CASE WHEN lastname IS NULL THEN '' ELSE lastname +' ' END)  
                            FROM tblEmpDetails WHERE id  = @id")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With intraTransfer
                For Each r As DataRow In SQL.DBDT.Rows
                    If Not IsDBNull(r(0)) Then

                        .userxxx = r(0)
                    Else
                        .userxxx = r(0)
                    End If
                Next
            End With
        Else

        End If
    End Sub

    Public Sub loadProductforTransfer(lv As ListView, tbx1 As TextBox, tbx2 As TextBox, aid As String)
        SQL.AddParam("@aId", aid)
        SQL.AddParam("@id", tbx1.Text)
        SQL.AddParam("@des", "%" & tbx2.Text & "%")

        SQL.ExecQueryDT("
                             SELECT TOP 50
        
                           t0.id,t1.id AS GoodID,t1.goodDes,t0.areaId ,t0.batchId ,t0.locationId,t2.batchName ,t3.name,t0.qty ,t0.Expirationdate
                           FROM tblProducts t0
                           inner join tblGoods t1 on t0.goodId = t1.id
                           inner join tblBatches t2 on t0.batchId = t2.id
                           inner join tblLocations t3 on t0.locationId = t3.id
                           where t0.qty>0 
                           AND (t0.areaId = @aId)
                           AND (t1.goodDes LIKE @des or @des='')
                           AND (t1.id like '%'+ @id +'%' or @id ='')
                           ORDER BY t1.id ASC")

        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                    .SubItems.Add(r(7))
                    .SubItems.Add(r(8))

                    If r.IsNull(9) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(9))
                    End If
                End With
            Next
        End If
    End Sub

    Public Sub selectLVLocationsIntraTransfer(lv As ListView, tbx As TextBox, id As String)
        SQL.AddParam("@flter", tbx.Text & "%")
        SQL.AddParam("@aId", newHome.areaId)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT id, [name], remarks FROM tblLocations
                        WHERE areaId = @aId AND [name] LIKE @flter
                        AND id <> @id;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub

    Public Sub addIntraTransfer(transId As String, transDate As Date, rmks As String, rmks1 As String, encBy As String, area As String, approvedBy As String)
        SQL.AddParam("@transid", transId)
        SQL.AddParam("@transDate", transDate)
        SQL.AddParam("@rmks", rmks)
        SQL.AddParam("@rmks1", rmks1)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@areaId", area)
        SQL.AddParam("@apprvBy", approvedBy)
        SQL.ExecQueryDT("INSERT INTO tblIntraTransferHeaders
                        (transferId, dateTransfer, remarks,remarks_1,encBy,encDate,areaId,approvedBy)
                        VALUES
                        (@transid, @transDate, @rmks, @rmks1, @encBy, GETDATE(),@areaId,@apprvBy)
                        SELECT SCOPE_IDENTITY();")
        If SQL.HasException(True) Then Exit Sub

        Dim entryID As String = ""

        If SQL.RecordCountDT > 0 Then

            For Each r As DataRow In SQL.DBDT.Rows
                entryID = r(0)
            Next

            With intraTransfer.dgvTransferGoods 'adding the datagridview
                For i As Integer = 0 To .Rows.Count - 2 Step +1
                    SQL.AddParam("@transID", entryID)
                    SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
                    SQL.AddParam("@gId", .Rows(i).Cells(5).Value.ToString)
                    SQL.AddParam("@bId", .Rows(i).Cells(1).Value.ToString)
                    SQL.AddParam("@lF", .Rows(i).Cells(2).Value.ToString)
                    SQL.AddParam("@lT", .Rows(i).Cells(3).Value.ToString)
                    SQL.AddParam("@aId", .Rows(i).Cells(4).Value.ToString)
                    SQL.AddParam("@qty", .Rows(i).Cells(10).Value.ToString)
                    SQL.AddParam("@exdate", .Rows(i).Cells(11).Value.ToString)
                    SQL.ExecQueryDT("INSERT INTO tblIntraTransferDetails
                            (transId, pId, gId, bId, lIdFrom, lIdTo, aId, qty,Expirationdate)
                            VALUES
                            (@transID, @pId, @gId, @bId, @lF, @lT, @aId, @qty,@exdate);")
                    If SQL.HasException(True) Then Exit Sub
                Next
            End With

            SQL.AddParam("@entryId", entryID)
            SQL.ExecQueryDT("SELECT * FROM tblIntraTransferDetails
            WHERE transId = @entryId;")
            If SQL.HasException(True) Then Exit Sub

            If SQL.RecordCountDT <> 0 Then
                Dim pId As String = ""
                Dim gId As String = ""
                Dim bId As String = ""
                Dim lIdF As String = ""
                Dim lIdT As String = ""
                Dim aId As String = ""
                Dim xdate As String = ""
                Dim qty As Double

                For Each r As DataRow In SQL.DBDT.Rows
                    pId = r("pId")
                    gId = r("gId")
                    bId = r("bId")
                    lIdF = r("lIdFrom")
                    lIdT = r("lIdTo")
                    aId = r("aId")
                    qty = r("qty")
                    xdate = r("Expirationdate")

                    SQL.AddParam("@pId", pId)
                    SQL.AddParam("@qty", qty)


                    SQL.ExecQueryDT("UPDATE tblProducts
                                    SET qty = qty - @qty
                                    WHERE id = @pId;")
                    If SQL.HasException(True) Then Exit Sub
                    SQL.AddParam("@gId", gId)
                    SQL.AddParam("@bId", bId)
                    SQL.AddParam("@lIdT", lIdT)
                    SQL.AddParam("@aId", aId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)

                    SQL.ExecQueryDT("SELECT * FROM tblProducts
                                    WHERE goodId = @gId
                                    AND areaId = @aId
                                    AND batchId = @bId
                                    AND locationId = @lIdT
                                    AND Expirationdate = @exdate")

                    If SQL.HasException(True) Then Exit Sub
                    If SQL.RecordCountDT > 0 Then 'already exists
                        'Dim prodId As String = ""
                        'For Each row As DataRow In SQL.DBDT.Rows
                        '    prodId = r("id")
                        'Next
                        SQL.AddParam("@gId", gId)
                        SQL.AddParam("@aId", aId)
                        SQL.AddParam("@bId", bId)
                        SQL.AddParam("@lId", lIdT)
                        SQL.AddParam("@qty", qty)
                        SQL.AddParam("@exdate", xdate)

                        SQL.ExecQueryDT("UPDATE tblProducts
                                        SET qty = qty + @qty
                                        WHERE goodId = @gId
                                        AND areaId = @aId
                                        AND batchId = @bId
                                        AND locationId = @lId
                                        AND Expirationdate = @exdate")

                        If SQL.HasException(True) Then Exit Sub
                    Else 'not exists
                        SQL.AddParam("@gId", gId)
                        SQL.AddParam("@aId", aId)
                        SQL.AddParam("@bId", bId)
                        SQL.AddParam("@lId", lIdT)
                        SQL.AddParam("@qty", qty)
                        SQL.AddParam("@exdate", xdate)
                        SQL.ExecQueryDT("INSERT INTO tblProducts
                                        (goodId,areaId,batchId,locationId,qty,Expirationdate)
                                        VALUES
                                        (@gId,@aId,@bId,@lId,@qty,@exdate)")

                        If SQL.HasException(True) Then Exit Sub
                    End If
                Next
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        MessageBox.Show("Transaction Successfully Recorded.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        With intraTransfer
            .dgvTransferGoods.Enabled = False
            .dgvTransferGoods.AllowUserToAddRows = False
            .btnAdd.Text = "Add New Entry"
            .btnAdd.Select()
            .tbxEntryNumber.Text = entryID
        End With
    End Sub

    Public Sub validateDocumentNumberIntraTransfer(docNum As String, area As String)
        SQL.AddParam("@docNum", docNum)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblIntraTransferHeaders WHERE transferId = @docNum AND areaId = @area ;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            If (MessageBox.Show("Document" & " #" & docNum & " already exists. " & vbCrLf & vbCrLf & "Click OK if you wish to record the transaction using a duplicate Document #." & vbCrLf & "Click CANCEL if you do not want to record the transaction.",
                                "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.Cancel) Then
                With intraTransfer
                    .tbxTransferId.Text = ""
                    .tbxTransferId.Select()
                End With
                Exit Sub
            Else

            End If
        Else
            Exit Sub
        End If

    End Sub

    Public Sub loadIntraTransRegister(lv As ListView, trnsferId As String, entryId As String, date1 As Date, date2 As Date, area As String, roleId As String, areaName As String)
        'supervisor roleId 2016 change to roleId 20017 
        'and manager roleId 20018 change to 20016 
        'and encoder roleId 20019 change to 20018 
        If roleId = "20016" Or roleId = "20016" Or roleId = "20020" Then
            SQL.AddParam("@strDate", date1)
            SQL.AddParam("@stpDate", date2)
            SQL.AddParam("@entryId", entryId & "%")
            SQL.AddParam("@trnsferId", trnsferId & "%")
            SQL.AddParam("@areaName", areaName & "%")
            SQL.ExecQueryDT("SELECT
                                '0' as dummy,
                                T0.dateTransfer
                                ,T0.id
                                ,T0.transferId
                                ,T0.remarks_1
                                ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
                                ,CONVERT(DATE, T0.encDate)
                                ,T0.remarks
                                FROM tblIntraTransferHeaders T0
                                JOIN tblEmpDetails T1 ON T1.id = T0.encBy
                        WHERE T0.dateTransfer BETWEEN @strDate AND @stpDate
                        AND T0.transferId LIKE @trnsferId
                        AND T0.id LIKE @entryId
                        AND (SELECT area FROM tblAreas WHERE id = T0.areaId) LIKE @areaName
                        ORDER BY T0.encDate;")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(4))
                        .SubItems.Add(r(5))
                        .SubItems.Add(r(6))
                        .SubItems.Add(r(7))
                    End With
                Next
            End If
        Else
            SQL.AddParam("@strDate", date1)
            SQL.AddParam("@stpDate", date2)
            SQL.AddParam("@aId", area)
            SQL.AddParam("@entryId", entryId & "%")
            SQL.AddParam("@trnsferId", trnsferId & "%")
            SQL.ExecQueryDT("SELECT
                        '0' as dummy,
                        T0.dateTransfer
                        ,T0.id
                        ,T0.transferId
                        ,T0.remarks_1
                        ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
                        ,CONVERT(DATE, T0.encDate)
                        ,T0.remarks
                        FROM tblIntraTransferHeaders T0
                        JOIN tblEmpDetails T1 ON T1.id = T0.encBy
                        WHERE T0.dateTransfer BETWEEN @strDate AND @stpDate
                        AND T0.transferId LIKE @trnsferId
                        AND T0.id LIKE @entryId
                        AND T0.areaId = @aId
                        ORDER BY T0.encDate;")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(4))
                        .SubItems.Add(r(5))
                        .SubItems.Add(r(6))
                        .SubItems.Add(r(7))
                    End With
                Next
            End If
        End If
    End Sub

    Public Sub fetchIntraTransferDetails(lv As ListView, id As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@aId", newHome.areaId)
        SQL.ExecQueryDT("SELECT '' as dummy, T0.gId,T1.goodDes,T2.batchName,T3.name,T4.name,T0.qty,T0.Expirationdate
                            FROM tblIntraTransferDetails T0
                            INNER JOIN tblGoods T1 ON T0.gId=T1.id
                            INNER JOIN tblBatches T2 ON T0.bId=T2.id
                            INNER JOIN tblLocations T3 ON T0.lIdFrom=T3.id
                            INNER JOIN tblLocations T4 ON T0.lIdTo=T4.id
                        WHERE transId = @id ORDER BY gId ASC;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                    If r.IsNull(7) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(7))
                    End If
                End With
            Next
        End If

        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT T0.id, T0.transferId, T0.dateTransfer, T0.remarks, T0.remarks_1,
                            (SELECT DISTINCT 
                                    CONCAT(COALESCE([name]+ ' ',''),
                                    COALESCE(lastname+ ' ',''),
                                    COALESCE(middlename,''))  
                                FROM tblEmpDetails WHERE id = T0.approvedBy)
                            FROM tblIntraTransferHeaders T0
                        WHERE T0.id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With intraTransferDetails
                    If Not IsDBNull(r(0)) Then
                        .tbxEntryNumber.Text = r(0)
                    Else
                        .tbxEntryNumber.Text = ""
                    End If

                    If Not IsDBNull(r(1)) Then
                        .tbxTransferId.Text = r(1)
                    Else
                        .tbxTransferId.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .dtpTransferDate.Value = r(2)
                    Else
                        .dtpTransferDate.Value = Date.Now
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxUser.Text = r(3)
                    Else
                        .tbxUser.Text = ""
                    End If

                    If Not IsDBNull(r(4)) Then
                        .tbxRemarks.Text = r(4)
                    Else
                        .tbxRemarks.Text = ""
                    End If
                    If Not IsDBNull(r(5)) Then
                        .tbxApprovedBy.Text = r(5)
                    Else
                        .tbxApprovedBy.Text = ""
                    End If
                End With
            Next
        Else

        End If
    End Sub

    Public Function validateIDinGoodsAutoBuild(id As String)
        Dim bool As Boolean = False
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("select * from tblGoods WHERE id like '%'+ @id +'%'")
        If SQL.HasException(True) Then Return Nothing

        If SQL.RecordCountDT > 0 Then
            bool = True
        Else
            bool = False
        End If

        Return bool
    End Function

    Public Sub addNewGoodsAutoBuild(tbx As String, dateEnc As String, encBy As String, apprv As String)
        Dim status As Integer = 1
        Dim defaultleadtime As Integer = 40
        Dim defaultleadtimeallowance As Integer = 40 * 0.25
        SQL.AddParam("@tbx", tbx)
        SQL.AddParam("@date", dateEnc)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@apprv", apprv)
        SQL.ExecQueryDT("INSERT INTO tblGoodsABHeaders
                        (encDate,encby,[date], tag, apprvBy)
                        VALUES
                        (@date, @encBy, GETDATE(), @tbx, @apprv)
                        SELECT SCOPE_IDENTITY();")
        If SQL.HasException(True) Then Exit Sub
        Dim transId As String = ""
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transId = r(0)
            Next
            With goodsAutoBuild.dgvGoods

                For i As Integer = 0 To .Rows.Count - 2 Step +1
                    SQL.AddParam("@transid", transId)
                    SQL.AddParam("@id", .Rows(i).Cells(0).Value.ToString)
                    SQL.AddParam("@des", .Rows(i).Cells(1).Value.ToString)

                    SQL.ExecQueryDT("INSERT INTO tblGoodsABDetails
                                (transId,goodId,goodDes)
                                VALUES	
                                (@transid,@id,@des);")
                    If SQL.HasException(True) Then Exit Sub

                    SQL.AddParam("@id", .Rows(i).Cells(0).Value.ToString)
                    SQL.AddParam("@des", .Rows(i).Cells(1).Value.ToString)
                    SQL.AddParam("@leadtime", defaultleadtime)
                    SQL.AddParam("@leadtimeallowance", defaultleadtimeallowance)
                    SQL.AddParam("@status", status)

                    SQL.ExecQueryDT("INSERT INTO tblGoods
                                (id,goodDes,LeadTime,LeadTimeAllowance,Status)
                                VALUES
                                (@id,@des,@leadtime,@leadtimeallowance,@status);")
                    If SQL.HasException(True) Then Exit Sub

                Next
            End With

            With goodsAutoBuild
                .tbxEntry.Text = transId
                .btnRecord.Text = "Add New"
                .lblErr.Visible = True
                .lblErr.ForeColor = Color.Green
                .lblErr.Text = "Added Successfully"
                .dgvGoods.Enabled = False
                .dgvGoods.AllowUserToAddRows = False
                .btnRecord.Select()
            End With
        Else
            Exit Sub
        End If
    End Sub

    Public Sub loadGoodsListLV(lv As ListView, tbx1 As TextBox, tbx2 As TextBox)

        SQL.AddParam("@id", tbx1.Text + "%")
        SQL.AddParam("@des", "%" & tbx2.Text & "%")
        SQL.ExecQueryDT("SELECT top 500 * FROM tblGoods
                        WHERE id LIKE @id
                        AND gooddes LIKE @des
                        ORDER BY id")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                End With
            Next
        End If



        SQL.ExecQueryDT("SELECT count(goodDes) FROM tblGoods")
        If SQL.HasException(True) Then Exit Sub
        With goodsList
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .Label3.Text = r(0).ToString + " TOTAL LIST OF GOODS"
                Next
            End If
        End With



    End Sub

    Public Sub loadGoodsRegisterHeader(lv As ListView, entryId As TextBox, strDate As Date, stpDate As Date)
        SQL.AddParam("@id", entryId.Text + "%")
        SQL.AddParam("@str", strDate)
        SQL.AddParam("@stp", stpDate)
        SQL.ExecQueryDT("SELECT T0.id, T0.encDate,
                            (SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
                                FROM tblEmpDetails WHERE id = T0.encBy), T0.tag, T0.[date]
                        FROM tblGoodsABHeaders T0
                        WHERE encDate BETWEEN @str AND @stp AND id LIKE @id
                        ORDER BY encDate;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2)) 'gi comment naku ni kay naay error
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                End With
            Next
        End If
    End Sub

    Public Sub loadGoodsRegisterDetails(lv As ListView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT * FROM tblGoodsABDetails WHERE transId = @id;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(2))
                With itemNew
                    If Not IsDBNull(r(3)) Then
                        .SubItems.Add(r(3))
                    Else
                        .SubItems.Add("")
                    End If
                End With
            Next
        End If

        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT T0.id, T0.encDate,
                            (SELECT CONCAT(COALESCE([name]+ '  ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
                            FROM tblEmpDetails WHERE id = T0.encBy),
                            T0.tag,
                             (SELECT CONCAT(COALESCE([name]+ '  ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
                                FROM tblEmpDetails WHERE id = T0.apprvBy)
                        FROM tblGoodsABHeaders T0
                        WHERE id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With goodsRegisterDetails
                    If Not IsDBNull(r(0)) Then
                        .tbxEntry.Text = r(0)
                    Else
                        .tbxEntry.Text = ""
                    End If

                    If Not IsDBNull(r(1)) Then
                        .dtpDateEnc.Value = r(1)
                    Else
                        .dtpDateEnc.Value = Date.Now
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxRemarks.Text = r(3)
                    Else
                        .tbxRemarks.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxUser.Text = r(2)
                    Else
                        .tbxUser.Text = ""
                    End If

                    If Not IsDBNull(r(4)) Then
                        .tbxApproved.Text = r(4)
                    Else
                        .tbxApproved.Text = ""
                    End If
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub loadLV_selectABProducts(lv As ListView, id As TextBox, des As TextBox)
        SQL.AddParam("@id", id.Text)
        SQL.AddParam("@des", "%" & des.Text & "%")
        SQL.ExecQueryDT("SELECT TOP 100 '0' as dummy,id,goodDes FROM tblGoods
                        WHERE (id like '%'+ @id +'%' or @id ='')
                        AND (goodDes LIKE @des or @des ='')  AND (Status = 1) ")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub

    Public Sub loadLV_selectABBatch(lv As ListView, id As TextBox, des As TextBox)
        SQL.AddParam("@id", id.Text + "%")
        SQL.AddParam("@des", des.Text + "%")
        SQL.ExecQueryDT("SELECT * FROM tblBatches
                        WHERE batchName LIKE @id
                        AND remarks LIKE @des;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub

    Public Sub loadLV_selectABLocations(lv As ListView, id As TextBox, des As TextBox)
        SQL.AddParam("@id", id.Text + "%")
        SQL.AddParam("@des", des.Text + "%")
        SQL.AddParam("@area", newHome.areaId)
        SQL.ExecQueryDT("SELECT * FROM tblLocations
                        WHERE [name] LIKE @id
                        AND remarks LIKE @des
                        AND areaId = @area;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                End With
            Next
        End If
    End Sub

    Public Sub addProductAutoBuild(docNum As String, transDate As Date, rmks As String, rmks1 As String, encBy As String, encArea As String, approvedId As String)
        Dim transId As String = ""
        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double = 0 'Integer change to Double for decimal values
        Dim exdate As String = ""


        SQL.AddParam("@docNum", docNum)
        SQL.AddParam("@transDate", transDate)
        SQL.AddParam("@rmks", rmks)
        SQL.AddParam("@rmks1", rmks1)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@aprvBy", approvedId)
        SQL.AddParam("@areaid", encArea)
        SQL.ExecQueryDT("INSERT INTO tblProductsABHeaders
                        (docNum,transDate,rmks,rmks_1,encDate,encBy,approvedBy,areaId)
                        VALUES
                        (@docNum,@transDate,@rmks,@rmks1,GETDATE(),@encBy,@aprvBy,@areaid)
                        SELECT SCOPE_IDENTITY();")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transId = r(0)
            Next
        Else
        End If

        With inventoryAutoBuild.dgvInventory
            For i As Integer = 0 To .Rows.Count - 2 Step +1
                SQL.AddParam("@transId", transId)
                SQL.AddParam("@gId", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@bId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@lId", .Rows(i).Cells(2).Value.ToString)
                SQL.AddParam("@aId", encArea)
                SQL.AddParam("@qty", .Rows(i).Cells(6).Value.ToString)
                SQL.AddParam("@Expiration", .Rows(i).Cells(7).Value.ToString)


                SQL.ExecQueryDT("INSERT INTO tblProductsABDetails
                                (transId, goodId, batchId,locationId,areaId, qty,ExpirationDate)
                                VALUES
                                (@transId, @gId, @bId, @lId, @aId, @qty,@Expiration);")
                If SQL.HasException(True) Then Exit Sub

            Next
        End With

        SQL.AddParam("@transId", transId)
        SQL.ExecQueryDT("SELECT goodId, batchId,locationId,areaId,qty,Expirationdate FROM tblProductsABDetails WHERE transId = @transId;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                goodId = r("goodId")
                batchId = r("batchId")
                locId = r("locationId")
                areaId = r("areaId")
                qty = r("qty")
                exdate = r("Expirationdate")

                SQL.AddParam("@gId", goodId)
                SQL.AddParam("@bId", batchId)
                SQL.AddParam("@lId", locId)
                SQL.AddParam("@aId", areaId)
                SQL.AddParam("@qty", qty)
                SQL.AddParam("@Expirationdate", exdate)

                SQL.ExecQueryDT("SELECT * FROM tblProducts
                                    WHERE goodId = @gId
                                    AND areaId = @aId
                                    AND batchId = @bId
                                    AND locationId = @lId
                                    AND Expirationdate = @Expirationdate")
                If SQL.HasException(True) Then Exit Sub

                If SQL.RecordCountDT > 0 Then
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@aId", areaId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", exdate)
                    SQL.ExecQueryDT("UPDATE tblProducts
								set qty = qty + @qty
                                WHERE goodId = @gId
								AND areaId = @aId
								AND batchId = @bId
								AND locationId = @lId
                                AND Expirationdate=@exdate;")
                    If SQL.HasException(True) Then Exit Sub
                Else
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@aId", areaId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", exdate)
                    SQL.ExecQueryDT("INSERT INTO tblProducts
								(goodId,areaId,batchId,locationId,qty,Expirationdate)
								VALUES
								(@gId,@aId,@bId,@lId,@qty,@exdate);")
                    If SQL.HasException(True) Then Exit Sub
                End If

            Next
        End If

        With inventoryAutoBuild
            .dgvInventory.Enabled = False
            .dgvInventory.AllowUserToAddRows = False
            .tbxEntryNumber.Text = transId
            .btnSave.Select()
            .btnSave.Text = "Add New Entry"
            .lblErr.Visible = True
            .lblErr.ForeColor = Color.Green
            .lblErr.Text = "Successfully Added."
        End With
    End Sub

    Public Sub loadAutoBuildProductsRegister(lv As ListView, tbxId As TextBox, tbxDocNo As TextBox, strDate As Date, stpDate As Date, area As String, roleId As String, areaName As String)
        'supervisor roleId 2016 change to roleId 20017 
        'and manager roleId 20018 change to 20016 
        'and encoder roleId 20019 change to 20018 
        If roleId = "20016" Or roleId = "20016" Or roleId = "20020" Then
            SQL.AddParam("@id", tbxId.Text + "%")
            SQL.AddParam("@docNum", tbxDocNo.Text + "%")
            SQL.AddParam("@str", strDate)
            SQL.AddParam("@stp", stpDate)
            SQL.AddParam("@area", area)
            SQL.AddParam("@areaName", areaName & "%")
            SQL.ExecQueryDT("SELECT 
                            '0' as dummy
                            ,T0.transDate
                            , T0.id
                            ,T0.docNum
                            , T0.rmks
                            ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
                            ,CONVERT(DATE, T0.encDate)
                            ,(SELECT CONCAT(COALESCE([name]+ '  ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  FROM tblEmpDetails WHERE id = T0.encBy)
                            FROM tblProductsABHeaders T0
                            JOIN tblEmpDetails T1
                            ON T1.id = T0.encBy
                        WHERE T0.id LIKE @id AND T0.docNum LIKE @docNum
                        AND T0.transDate BETWEEN @str AND @stp
                        AND (SELECT area FROM tblAreas WHERE id = T0.areaId) LIKE @areaName
                        ORDER BY T0.encDate;")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        If Not IsDBNull(r(1)) Then
                            .SubItems.Add(r(1))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(2)) Then
                            .SubItems.Add(r(2))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(3)) Then
                            .SubItems.Add(r(3))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(4)) Then
                            .SubItems.Add(r(4))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(5)) Then
                            .SubItems.Add(r(5))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(6)) Then
                            .SubItems.Add(r(6))
                        Else
                            .SubItems.Add("")
                        End If


                        If Not IsDBNull(r(7)) Then
                            .SubItems.Add(r(7))
                        Else
                            .SubItems.Add("")
                        End If
                    End With
                Next
            End If
        Else
            SQL.AddParam("@id", tbxId.Text + "%")
            SQL.AddParam("@docNum", tbxDocNo.Text + "%")
            SQL.AddParam("@str", strDate)
            SQL.AddParam("@stp", stpDate)
            SQL.AddParam("@area", area)
            SQL.ExecQueryDT("SELECT 
                            '0' as dummy,
                            T0.transDate
                            , T0.id
                            ,T0.docNum
                            , T0.rmks
                            ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
                            ,CONVERT(DATE, T0.encDate)
                            ,(SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ '',''),COALESCE(middlename,''))  FROM tblEmpDetails WHERE id = T0.encBy)
                            FROM tblProductsABHeaders T0
                            JOIN tblEmpDetails T1
                            ON T1.id = T0.encBy
                        WHERE T0.id LIKE @id AND T0.docNum LIKE @docNum
                        AND T0.transDate BETWEEN @str AND @stp
                        AND T0.areaId = @area
                        ORDER BY T0.encDate;")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        If Not IsDBNull(r(1)) Then
                            .SubItems.Add(r(1))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(2)) Then
                            .SubItems.Add(r(2))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(3)) Then
                            .SubItems.Add(r(3))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(4)) Then
                            .SubItems.Add(r(4))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(5)) Then
                            .SubItems.Add(r(5))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(6)) Then
                            .SubItems.Add(r(6))
                        Else
                            .SubItems.Add("")
                        End If

                        If Not IsDBNull(r(7)) Then
                            .SubItems.Add(r(7))
                        Else
                            .SubItems.Add("")
                        End If
                    End With
                Next
            End If
        End If

    End Sub

    Public Sub fetchLoadProductAutoBuild(lv As ListView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT '0' as dummy,T0.goodId,T1.goodDes,T2.batchName,T3.name,T0.qty,T0.ExpirationDate
                            FROM tblProductsABDetails T0
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            INNER JOIN tblBatches T2 ON T0.batchId=T2.id
                            INNER JOIN tblLocations T3 ON T0.locationId=T3.id
                            WHERE T0.transId =@id
                            ORDER BY T0.goodId ASC")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(If(r(1) IsNot DBNull.Value, r(1).ToString(), ""))
                    .SubItems.Add(If(r(2) IsNot DBNull.Value, r(2).ToString(), ""))
                    .SubItems.Add(If(r(3) IsNot DBNull.Value, r(3).ToString(), ""))
                    .SubItems.Add(If(r(4) IsNot DBNull.Value, r(4).ToString(), ""))
                    .SubItems.Add(If(r(5) IsNot DBNull.Value, r(5).ToString(), ""))
                    .SubItems.Add(If(r(6) IsNot DBNull.Value, r(6).ToString(), ""))
                End With
            Next
        End If

        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT
                            T0.transDate
                            ,T0.id
                            ,T0.docNum
                            ,T0.rmks
                            ,T0.rmks_1
                            ,(SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  FROM tblEmpDetails WHERE id = T0.encBy)
                            ,(SELECT CONCAT(COALESCE([name]+ ' ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  FROM tblEmpDetails WHERE id = T0.approvedBy)
                            ,T0.encDate  
                        FROM tblProductsABHeaders T0
                        WHERE T0.id = @id;")
        If SQL.HasException(True) Then Exit Sub
        With inventoryGoodsDetails
            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    If Not IsDBNull(r(0)) Then
                        .dtpDate.Value = r(0)
                    Else
                        .dtpDate.Value = Date.Now
                    End If

                    If Not IsDBNull(r(1)) Then
                        .tbxEntryNumber.Text = r(1)
                    Else
                        .tbxEntryNumber.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDocNum.Text = r(2)
                    Else
                        .tbxDocNum.Text = ""
                    End If

                    If Not IsDBNull(r(3)) Then
                        .tbxRemarks.Text = r(3)
                    Else
                        .tbxRemarks.Text = ""
                    End If

                    If Not IsDBNull(r(4)) Then
                        .cbxOwnership.Text = r(4)
                    Else
                        .cbxOwnership.Text = ""
                    End If

                    If Not IsDBNull(r(5)) Then
                        .tbxUser.Text = r(5)
                    Else
                        .tbxUser.Text = ""
                    End If

                    If Not IsDBNull(r(6)) Then
                        .tbxApprvdBy.Text = r(6)
                    Else
                        .tbxApprvdBy.Text = ""
                    End If
                Next
            Else
                Exit Sub
            End If
        End With
    End Sub

    'Public Sub loadInventoryIGLR(cbx As CheckedListBox, tbx As TextBox, tbx1 As TextBox)
    '    'Public Sub loadInventoryIGLRr(tbx As TextBox, tbx1 As TextBox)
    '    SQL.AddParam("@des", tbx.Text & "%")
    '    SQL.AddParam("@id", tbx1.Text & "%")
    '    SQL.ExecQueryDT("SELECT TOP 10 CONCAT('(',id,')',goodDes) FROM tblGoods WHERE (goodDes LIKE @des AND id LIKE @id);")
    '    If SQL.HasException(True) Then Exit Sub

    '    'If SQL.RecordCountDT > 0 Then
    '    '    For Each r As DataRow In SQL.DBDT.Rows
    '    '        cbx.Items.Add(r(0))
    '    '    Next
    '    'Else
    '    '    Exit Sub
    '    'End If
    'End Sub


    Public Sub loadInventoryIGLR(cbx As ListView, tbx As String, tbx1 As String)
        cbx.Items.Clear()
        SQL.AddParam("@des", "%" & tbx & "%")
        SQL.AddParam("@id", tbx1)
        SQL.ExecQueryDT("SELECT TOP 50 *
        FROM tblGoods T0
        WHERE (T0.goodDes LIKE @des OR @des = '')
        AND (CAST(T0.id AS VARCHAR(50)) LIKE @id OR @id = '' OR T0.id = CAST(@id AS INT))")
        If SQL.HasException(True) Then Exit Sub
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = cbx.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub loadInventoryIGLRr(cbx As ListView, tbx As String, tbx1 As String)
        cbx.Items.Clear()
        SQL.AddParam("@des", tbx & "%")
        SQL.AddParam("@id", tbx1)
        SQL.ExecQueryDT("SELECT TOP 50 *
        FROM tblGoods T0
        WHERE (T0.goodDes LIKE @des OR @des = '')
        AND (T0.id = @id or @id='')")
        If SQL.HasException(True) Then Exit Sub
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = cbx.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub



    'TO EXECUTE. SCRIPT CUSTOMIZATION AS SEPT 22 2022 THURSDAY 9:47AM
    'Inventory -- Count Inventory -- New Count(Data Grid View)

    Public Sub loadDGVInventoryAdjust(t1 As String, t2 As TextBox, t3 As TextBox, area As String, Optional query As String = "")

        If query = "" Then
            SQL.AddParam("@gid", t1)
            SQL.AddParam("@des", "%" & t2.Text & "%")
            SQL.AddParam("@loc", "%" & t3.Text & "%")
            SQL.AddParam("@areaId", area)
            SQL.ExecQueryDT("Select TOP 100
                             t0.id
                             ,t0.goodId AS 'PRODUCT ID'
                             ,t1.goodDes AS 'DESCRIPTION'
                             ,t3.name AS 'LOCATION'
                             ,t2.batchName AS 'BATCH'
                                ,isnull(t0.Expirationdate,'N/A') as 'EXPIRATION DATE'
                             ,1 * T0.qty AS 'ON HAND'
                             ,1 * T0.qty AS 'COUNTED'
                             ,1 * (T0.qty - T0.qty) AS 'DIFFERENCE'
                             from tblProducts t0
                             inner join tblGoods t1 on t0.goodId = t1.id
                             inner join tblBatches t2 on t0.batchId = t2.id
                             inner join tblLocations t3 on t0.locationId = t3.id 
                            WHERE (t0.goodId LIKE '%' + @gid + '%' )
                             AND (t1.goodDes LIKE @des)
                          AND (t0.areaId = @areaId)
                         AND (t3.name LIKE @loc)
                         AND (t1.Status = 1) 
                        ORDER by t0.goodId, CASE
                                    WHEN ISNUMERIC(name) = 1 THEN REPLICATE('0', 100 - LEN(name)) + name
                                    ELSE name 
        	end asc")
            If SQL.HasException(True) Then Exit Sub
        Else
            SQL.ExecQueryDT(query)
        End If

        With adjustInventory.dgvProducts
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT
            .ClearSelection()
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .Columns(0).Visible = False
            .Columns(1).HeaderText = "PRODUCT NUMBER"
            .Columns(2).HeaderText = "DESCRIPTION"
            .Columns(4).HeaderText = "BATCH"
            .Columns(3).HeaderText = "LOCATION"
            .Columns(5).HeaderText = "EXPIRATION DATE"
            .Columns(6).HeaderText = "ON HAND"
            .Columns(7).HeaderText = "COUNT"
            .Columns(8).HeaderText = "DIFFERENCE"
            .Columns(1).Width = 105
            .Columns(2).Width = 490
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 80
            .Columns(7).Width = 85
            .Columns(8).Width = 100
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            '.Columns(6).ReadOnly = True
            '.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
                .ColumnHeadersDefaultCellStyle.Font = New Font(adjustInventory.dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
                .DefaultCellStyle.Font = New Font(adjustInventory.dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Regular)
            Next
            .CurrentCell = Nothing
        End With
    End Sub


    Public Sub loadexdate(t1 As String, t2 As String, t3 As String, Optional query As String = "")

        If query = "" Then
            SQL.AddParam("@gid", t1)
            SQL.AddParam("@des", "%" & t2 & "%")
            SQL.AddParam("@loc", t3)
            SQL.ExecQueryDT("Select TOP 50
                                t0.id
                                ,t0.goodId AS 'PRODUCT ID'
                                ,t1.goodDes AS 'DESCRIPTION'
                                ,t2.batchName AS 'BATCH'
                                ,t3.name AS 'LOCATION'
                                ,t0.qty AS 'QUANTITY'
                                ,isnull(t0.Expirationdate,'N/A') as 'EXPIRATION DATE'
                                ,'0' AS 'DIFFERENCE'
                                from tblProducts t0
                                inner join tblGoods t1 on t0.goodId = t1.id
                                inner join tblBatches t2 on t0.batchId = t2.id
                                inner join tblLocations t3 on t0.locationId = t3.id 
                                WHERE (t0.goodId like '%'+ @gid +'%' or @gid ='')
                                AND (t1.goodDes LIKE @des or @des = '')
                                AND (t3.name = @loc or @loc = '') ORDER by t0.goodId ASC")
            If SQL.HasException(True) Then Exit Sub
        Else
            SQL.ExecQueryDT(query)
        End If

        With EditExpirationdate.dgexdate
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT
            .ClearSelection()
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .Columns(0).Visible = False
            .Columns(1).HeaderText = "PRODUCT NUMBER"
            .Columns(2).HeaderText = "DESCRIPTION"
            .Columns(3).HeaderText = "BATCH"
            .Columns(4).HeaderText = "LOCATION"
            .Columns(5).HeaderText = "QUANTITY"
            .Columns(6).HeaderText = "EXPIRATION DATE"
            .Columns(7).HeaderText = "DIFF"
            .Columns(1).Width = 100
            .Columns(2).Width = 500
            .Columns(3).Width = 100
            .Columns(4).Width = 100
            .Columns(5).Width = 100
            .Columns(6).Width = 130
            .Columns(7).Width = 0

            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True
            .Columns(5).ReadOnly = True
            .Columns(6).ReadOnly = False
            .Columns(7).Visible = False
            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .CurrentCell = Nothing
        End With
    End Sub



    ''QUERY TO EDIT AND EXEC AS OF SEPT 22 THURSDAY 10:04AM
    'Public Sub inventoryAdjustFilter(pdno As String, pdDes As String, pdLoc As String, ByVal dt As DataGridViewRow area As String)
    '    SQL.AddParam("@pdno", pdno & "%")
    '    SQL.AddParam("@pdDes", pdDes & "%")
    '    SQL.AddParam("@pdLoc", pdLoc & "%")
    '    SQL.AddParam("@areaId", area)
    '    loadDGVInventoryAdjust(area, "SELECT
    '                    T0.goodId AS 'PRODUCT ID'
    '                    ,(SELECT goodDes FROM tblGoods WHERE id = T0.goodId) AS 'DESCRIPTION'
    '                    ,(SELECT batchName FROM tblBatches WHERE id = T0.batchId) AS 'BATCH'
    '                    ,(SELECT [name] FROM tblLocations WHERE id = T0.locationId) AS 'LOCATION'
    '                    ,1 * T0.qty AS 'ON HAND'
    '                    ,1 * T0.qty AS 'COUNTED'
    '                    ,1 * (T0.qty - T0.qty) AS 'DIFFERENCE'
    '                    FROM tblProducts T0 WHERE T0.areaId = @areaId
    '                    AND ((SELECT goodDes FROM tblGoods WHERE id = T0.goodId) LIKE @pdDes
    '                    AND T0.goodId LIKE @pdno
    '                    AND (SELECT [name] FROM tblLocations WHERE id = T0.locationId) LIKE @pdLoc)")
    '    'ORDER BY [PRODUCT ID] ASC;")
    '    If SQL.HasException(True) Then Exit Sub
    'End Sub

    ''to try to exec if it is working (executed pero naay error)
    'Public Sub inventoryAdjustFilter(pdno As String, pdDes As String, pdLoc As String, area As String)
    '    SQL.AddParam("@pdno", pdno & "%")
    '    SQL.AddParam("@pdDes", pdDes & "%")
    '    SQL.AddParam("@pdLoc", pdLoc & "%")
    '    SQL.AddParam("@areaId", area)
    '    loadDGVInventoryAdjust(area, "EXEC [dbo].[sp_inventoryAdjustFilter] @pdno, @pdDes, @pdLoc, @areaId")
    '    If SQL.HasException(True) Then Exit Sub
    'End Sub

    Public Sub suggestPDNOinInventoryAdjust(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT id from tblGoods ORDER BY id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub suggestPDDesinInventoryAdjust(ByVal d As AutoCompleteStringCollection)
        SQL.ExecQueryDT("SELECT goodDes from tblGoods ORDER BY id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub suggestProdLocinInventoryAdjust(ByVal d As AutoCompleteStringCollection, area As String)
        SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT [name] FROM tbllocations WHERE areaId = @area ORDER BY [name];")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                d.Add(r(0)).ToString()
            Next
        End If
    End Sub

    Public Sub inventoryAdjustmentApproval(user As String, pass As String, area As String)
        '--NEW COUNT ADJUSTMENT FORM WITH CONFIRMATION FORM
        SQL.AddParam("@user", user)
        SQL.AddParam("@pass", pass)
        'SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
                        INNER JOIN tblEmpDetails T1 ON T0.empId = T1.id
                        INNER JOIN tblRoles T2 ON T1.role_id = T2.id
	                    WHERE T0.username = @user
                        AND (T0.pass = @pass)
	                    AND (T1.status = 1)
                        AND (T1.role_id IN (20016, 20017,20020));") '--approval procedure approves only by the supervisor or manager ---(T1.role_id IN (20016, 20017));") 
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With adjustInventory
                For Each r As DataRow In SQL.DBDT.Rows
                    .approvedID = r(0)
                Next
                newHome.Enabled = True
                adjustInventory.Enabled = True
                confirmationFormAdjustInventory.Dispose()
                .addDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            With confirmationFormAdjustInventory
                .tbxUser.Text = ""
                .tbxPass.Text = ""
                .lblError.Visible = False
                .tbxUser.Select()
                .tbxPass.Select() '--added script
            End With
            Exit Sub
        End If
        MessageBox.Show("Transaction Successfully Recorded.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub fetchUserLogInAdjustInventory(id)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT CONCAT(CASE WHEN name IS NULL THEN '' ELSE name +' ' END,
                            CASE WHEN middlename IS NULL THEN '' ELSE middlename +' ' END,
                            CASE WHEN lastname IS NULL THEN '' ELSE lastname +' ' END)
                            FROM tblEmpDetails WHERE id  = @id")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With adjustInventory
                For Each r As DataRow In SQL.DBDT.Rows
                    If Not IsDBNull(r(0)) Then
                        .tbxUser.Text = r(0)
                    Else
                        .tbxUser.Text = ""
                    End If
                Next
            End With
        Else

        End If
    End Sub


    Public Sub fetchreasoninadjustment(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select  t1.description from tblReasoncode t1 where t1.module='New Count'  order by t1.reasoncode ASC")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub

    Public Sub fetchreasonnewcount(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select t1.description from tblReasoncode t1 where t1.module='New Inventory'  order by t1.reasoncode ASC")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub


    Public Sub fetchreasonrec(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select t1.description from tblReasoncode t1 where t1.module='Receive Goods'  order by t1.reasoncode ASC")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub


    Public Sub fetchreasonrel(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select t1.description from tblReasoncode t1 where t1.module='Release Goods'  order by t1.reasoncode ASC")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub

    Public Sub fetchreasontran(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("select t1.description from tblReasoncode t1 where t1.module='Transfer Goods'  order by t1.reasoncode ASC")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub


    Public Sub fetchrAreaAdjustInventory(area)
        SQL.AddParam("@id", area)
        SQL.ExecQueryDT("select area from tblAreas where id = @id")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            With adjustInventory
                For Each r As DataRow In SQL.DBDT.Rows
                    If Not IsDBNull(r(0)) Then
                        .TextBox4.Text = r(0)
                    Else
                        .TextBox4.Text = ""
                    End If
                Next
            End With
        Else

        End If
    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub addAdjustInventory(docnum As String, rmks As String, user As String, trnsDate As Date, approvBy As String, encBy As String, area As String)
        Dim transID As String = ""
        Dim pId As String = ""
        Dim adjQty As Double

        SQL.AddParam("@docnum", docnum)
        SQL.AddParam("@rmks", rmks)
        SQL.AddParam("@user", user)
        SQL.AddParam("@trnsDate", trnsDate)
        SQL.AddParam("@approvBy", approvBy)
        SQL.AddParam("@encBy", encBy)
        SQL.AddParam("@areaId", area)
        'MessageBox.Show("start in header")
        SQL.ExecQueryDT("INSERT INTO tblInventoryAdjustHeaders
	                    (rmks,rmks1,userName,transDate,approvBy,encBy,encDate,areaId)
	                    VALUES
	                    (@rmks,@docnum,@user,@trnsDate,@approvBy,@encBY,GETDATE(),@areaId)
	                    SELECT SCOPE_IDENTITY()")

        If SQL.HasException(True) Then Exit Sub


        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                transID = r(0)
            Next

            With adjustInventory.dgvProducts
                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    If .Rows(i).Cells(8).Value <> 0 Then

                        SQL.AddParam("@transId", transID)
                        SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
                        SQL.AddParam("@counted", .Rows(i).Cells(6).Value.ToString)
                        SQL.AddParam("@onhand", .Rows(i).Cells(7).Value.ToString)
                        SQL.AddParam("@adjustQty", .Rows(i).Cells(8).Value.ToString)
                        SQL.AddParam("@exdate", .Rows(i).Cells(5).Value.ToString)

                        SQL.ExecQueryDT("INSERT INTO tblInventoryAdjustDetails
                                         (transId,pId,counted,onHand,adjustment,Expirationdate)
                                         VALUES
                                         (@transId,@pId,@counted,@onhand,@adjustQty,@exdate);")
                        If SQL.HasException(True) Then Exit Sub
                    End If
                    'MessageBox.Show("end in details")
                Next

                SQL.AddParam("@entryId", transID)
                SQL.ExecQueryDT("SELECT * FROM [dbo].[tblInventoryAdjustDetails] WHERE transid = @entryId;")
                If SQL.HasException(True) Then Exit Sub

                If SQL.RecordCountDT > 0 Then
                    For Each r As DataRow In SQL.DBDT.Rows
                        pId = r("pId")
                        adjQty = r("adjustment")

                        SQL.AddParam("@pId", pId)
                        SQL.AddParam("@adjQty", adjQty)

                        SQL.ExecQueryDT("UPDATE tblProducts
	                        SET qty = qty + @adjQty
	                        WHERE id = @pId;")

                        If SQL.HasException(True) Then Exit Sub

                    Next
                End If


            End With
        Else
            Exit Sub
        End If

        With adjustInventory
            .Refresh()
            .dgvProducts.Refresh()
            .tbxEntryNumber.Text = transID
            .dgvProducts.Enabled = True
            .btnSave.Text = "Add New"
            .btnSave.Select()
        End With
    End Sub






    'Public Sub updateAdjustInventory(dtg As DataGridView)

    '    Dim transID As String = ""
    '    Dim pId As String = ""
    '    Dim adjQty As String = Convert.ToDouble(adjQty)

    '    With adjustInventory.dgvProducts
    '        For i As Integer = 0 To .Rows.Count - 1 Step +1
    '            If .Rows(i).Cells(7).Value <> 0 Then
    '                MsgBox("new update phase")
    '                SQL.AddParam("@transId", transID)
    '                SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
    '                SQL.AddParam("@counted", .Rows(i).Cells(6).Value.ToString)
    '                SQL.AddParam("@onhand", .Rows(i).Cells(5).Value.ToString)
    '                SQL.AddParam("@adjustQty", .Rows(i).Cells(7).Value.ToString)

    '                SQL.ExecQueryDT("INSERT INTO tblInventoryAdjustDetails
    '                                     (transId,pId,counted,onHand,adjustment)
    '                                     VALUES
    '                                     (@transId,@pId,@counted,@onhand,@adjustQty);")
    '                If SQL.HasException(True) Then Exit Sub
    '            End If
    '        Next

    '    End With
    'End Sub



































    '------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub loadLVforInventoryAdjustNONManager(lv As ListView, tbxEntry As String, area As String, strDate As Date, stpDate As Date)
        SQL.AddParam("@entry", tbxEntry & "%")
        SQL.AddParam("@area", area)
        SQL.AddParam("@strDate", strDate)
        SQL.AddParam("@stpDate", stpDate)
        SQL.ExecQueryDT("SELECT 
                        '0' as dummy,
                        transdate
                        ,T0.id
                        ,T0.rmks1
                        ,T0.rmks
                        ,CONVERT(DATE, T0.encDate)
                        ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
                        ,T0.userName
                        ,T0.encBy
                        FROM [dbo].[tblInventoryAdjustHeaders] T0
                        WHERE areaid = @area
                        AND (T0.transDate BETWEEN @strDate AND @stpDate)
                        AND T0.id LIKE @entry
                        ORDER BY T0.encDate;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    If r.IsNull(3) Then
                        .SubItems.Add("n/a")
                    Else
                        .SubItems.Add(r(3))
                    End If
                    If r.IsNull(4) Then
                        .SubItems.Add("n/a")
                    Else
                        .SubItems.Add(r(4))
                    End If
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                    .SubItems.Add(r(7))
                    .SubItems.Add(r(8))
                End With
            Next
        End If
    End Sub


    Public Sub loadLVforInventoryAdjustManager(lv As ListView, tbxEntry As String, isAll As Boolean, area As String, strDate As Date, stpDate As Date)
        If isAll = True Then
            SQL.AddParam("@entry", tbxEntry & "%")
            SQL.AddParam("@strDate", strDate)
            SQL.AddParam("@encDate", stpDate)
            SQL.ExecQueryDT("
							select 
                            '0' as dummy
							,t0.transDate
							,t0.id
							,t0.rmks1
                            ,t0.rmks
							,t1.area
							,t0.encDate
							,t0.userName
                            ,t0.encBy
							from tblInventoryAdjustHeaders t0
							inner join tblAreas t1 on t0.areaId = t1.id
                            where t0.transDate Between @strDate and @encDate                        
                            AND t0.id LIKE @entry
							order by t0.encDate

                            ")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        If r.IsNull(3) Then
                            .SubItems.Add("n/a")
                        Else
                            .SubItems.Add(r(3))
                        End If
                        If r.IsNull(4) Then
                            .SubItems.Add("n/a")
                        Else
                            .SubItems.Add(r(4))
                        End If
                        .SubItems.Add(r(5))
                        .SubItems.Add(r(6))
                        .SubItems.Add(r(7))
                        .SubItems.Add(r(8))

                        'If r.IsNull(4) Then
                        '    .SubItems.Add("n/a")
                        'Else
                        '    .SubItems.Add(r(4))
                        'End If
                        '.SubItems.Add(r(3))
                        '.SubItems.Add(r(5))
                        '.SubItems.Add(r(6))
                        '.SubItems.Add(r(2))

                    End With
                Next
            End If
        Else
            SQL.AddParam("@entry", tbxEntry & "%")
            SQL.AddParam("@strDate", strDate)
            SQL.AddParam("@stpDate", stpDate)
            SQL.AddParam("@area", area)
            SQL.ExecQueryDT("select 
                                 '0' as dummy
							,t0.transDate
							,t0.id
							,t0.rmks
                            ,t0.rmks1
							,t1.area
							,t0.encDate
							,t0.userName
                            ,t0.encBy
							from tblInventoryAdjustHeaders t0
							inner join tblAreas t1 on t0.areaId = t1.id
                            where t0.transDate between @strDate and @stpDate
                            AND t0.id LIKE @entry
	                        and t0.areaId = @area
							order by t0.encDate")
            If SQL.HasException(True) Then Exit Sub
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))

                        If r.IsNull(4) Then
                            .SubItems.Add("n/a")
                        Else
                            .SubItems.Add(r(4))
                        End If
                        .SubItems.Add(r(3))
                        .SubItems.Add(r(5))
                        .SubItems.Add(r(6))
                        .SubItems.Add(r(7))
                        .SubItems.Add(r(8))
                    End With
                Next
            End If
        End If
    End Sub

    Public Sub fetchDatasInventoryAdjustDetails(lv As ListView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT '' as dummy,
                        (SELECT goodId FROM tblProducts WHERE id = T0.pId)
                        ,(SELECT (SELECT goodDes FROM tblGoods WHERE id = T1.goodId) FROM tblProducts T1 WHERE T1.id = T0.pId)
                        ,(SELECT (SELECT batchName FROM tblBatches WHERE id = T1.batchId) FROM tblProducts T1 WHERE T1.id = T0.pId)
                        ,(SELECT (SELECT [name] FROM tblLocations WHERE id = T1.locationId) FROM tblProducts T1 WHERE T1.id = T0.pId)
                        ,T0.Expirationdate
                        ,T0.counted
                        ,T0.onHand
                        ,T0.adjustment
                        FROM tblInventoryAdjustDetails T0
                        WHERE T0.transId = @id
                        ORDER BY 
                        (SELECT goodId FROM tblProducts WHERE id = T0.pId);")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    If r.IsNull(5) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(5))
                    End If
                    .SubItems.Add(r(6))
                    .SubItems.Add(r(7))
                    .SubItems.Add(r(8))

                End With
            Next
        End If

        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("SELECT T0.id, T0.rmks, T0.userName, T0.transDate,
                         (SELECT CONCAT(CASE WHEN name IS NULL THEN '' ELSE name +' ' END,
                                        CASE WHEN middlename IS NULL THEN '' ELSE middlename +' ' END,
                                        CASE WHEN lastname IS NULL THEN '' ELSE lastname +' ' END)  
                                    FROM tblEmpDetails WHERE id = T0.approvBy)
                        FROM tblInventoryAdjustHeaders T0 WHERE T0.id = @id;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With adjustInventoryRegisterDetails
                    .tbxEntryNum.Text = r(0)
                    .tbxRemarks.Text = r(1)
                    .dtpDate.Value = r(3)

                End With
            Next
        Else
            Exit Sub
        End If
    End Sub





    Public Sub addNewBatchMain(bCode As String, rmks As String, user As String)
        SQL.AddParam("@b", bCode)
        SQL.AddParam("@rmks", rmks)

        SQL.ExecQueryDT("SELECT * FROM tblBatches WHERE batchName = @b;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With addBatchMain
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxBatchCode.Text = ""
                .tbxBatchCode.Select()
            End With
        Else
            With addBatchMain
                .lblErr.Visible = False
            End With

            SQL.AddParam("@b", bCode.ToUpper())
            SQL.AddParam("@r", rmks)
            SQL.AddParam("@u", user)

            SQL.ExecQueryDT("INSERT INTO tblBatches
                (batchName, remarks, encBy, encDate)
                VALUES
                (@b,@r,@u,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Batch code successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With addBatchMain
                newHome.Enabled = True
                viewBatch.Enabled = True
                .Dispose()
                viewBatch.Select()
                viewBatch.Focus()
            End With
            With viewBatch
                .loadLV()
                .lvBatches.SelectedItems.Clear()
                .tbxFilter.Select()
            End With
        End If
    End Sub

    Public Sub addNewLocationMain(loc As String, rmks As String, area As String, userId As String)
        SQL.AddParam("@l", loc)
        SQL.AddParam("@a", area)
        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE [name] = @l AND areaId = @a;")
        If SQL.RecordCountDT > 0 Then
            With addLocationMain
                .tbxLocName.Text = ""
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxLocName.Select()
            End With
        Else
            With addLocationMain
                .lblErr.Visible = False
            End With

            SQL.AddParam("@l", loc.ToUpper())
            SQL.AddParam("@r", rmks)
            SQL.AddParam("@a", area)
            SQL.AddParam("@u", userId)
            SQL.ExecQueryDT("INSERT INTO tblLocations
                            ([name], remarks,encby,areaId,lastModified)
                            VALUES
                            (@l,@r,@u,@a,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Location successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With addLocationMain
                newHome.Enabled = True
                viewLocations.Enabled = True
                .Dispose()
                viewLocations.Focus()
                viewLocations.Select()
            End With
            With viewLocations
                .loadLV()
            End With
        End If
    End Sub

    Public Sub validateIntraTransferConfirmation(user As String, pass As String, area As String)
        SQL.AddParam("@user", user)
        SQL.AddParam("@pass", pass)
        SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
	                    INNER JOIN tblEmpDetails T1 ON T0.empId = T1.id
                        INNER JOIN tblRoles T2 ON T1.role_id=T2.id 
	                    WHERE (T0.username = @user) 
                        AND (T0.pass = @pass)
                        AND (T1.status = 1)
                        AND (T1.role_id IN (20016, 20017,20020));") '--approval procedure approves only by the supervisor or manager

        'OR (T1.role_id = 20016);")      AND (area_id = @area)

        'SQL.ExecQueryDT("SELECT T0.empId FROM tblUsers T0
        '             INNER JOIN tblEmpDetails T1 ON T1.id = T0.empId
        '             WHERE (T0.username = @user) 
        '                AND (T0.pass = @pass)
        '             AND (area_id = @area)
        '                AND (T1.status = 1)
        '                AND (T1.role_id = 20018) 
        '                OR (T1.role_id = 20016);")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With intraTransfer
                For Each r As DataRow In SQL.DBDT.Rows
                    .approvedId = r(0)
                Next
                newHome.Enabled = True
                .Enabled = True
                intraTransferConfirmation.Dispose()
                .addDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            With intraTransferConfirmation
                .tbxUser.Text = ""  '"@user"
                .tbxPass.Text = ""  '"@pass"
                .lblError.Visible = False
                .tbxUser.Select()
            End With
            Exit Sub
        End If
    End Sub

    Public Sub inventoryAutoBuildConfirmationQry(user As String, pass As String, area As String)
        SQL.AddParam("@user", user)
        SQL.AddParam("@pass", pass)
        SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
	                    INNER JOIN tblEmpDetails T1 ON T0.empId = T1.id
                        INNER JOIN tblRoles T2 ON T1.role_id = T2.id
	                    WHERE T0.username = @user 
                        AND (T0.pass = @pass)
                        AND (T1.status = 1)
                        AND (T1.role_id IN (20016, 20017,20020));") '--approval procedure approves only by the supervisor or manager and Superuser


        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With inventoryAutoBuild
                For Each r As DataRow In SQL.DBDT.Rows
                    .approvBy = r(0)
                Next
                newHome.Enabled = True
                .Enabled = True
                .Focus()
                .Select()
                inventoryAutoBuildConfirmation.Dispose()
                .addDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            With inventoryAutoBuildConfirmation
                .tbxUser.Text = ""
                .tbxPass.Text = ""
                .lblError.Visible = False
                .tbxUser.Select()
            End With
            Exit Sub
        End If
    End Sub

    Public Sub validateInventoryAutoBuildDocNum(docNum As String, area As String)
        SQL.AddParam("@docNo", docNum)
        SQL.AddParam("@areaId", area)

        SQL.ExecQueryDT("SELECT * FROM tblProductsABHeaders WHERE docNum = @docNo AND areaId = @areaId;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            If (MessageBox.Show("Document" & " #" & docNum & " has been previously used. " & vbCrLf & vbCrLf &
                                "Click OK if you wish to record the transaction using a duplicate " & "Document" & " #." & vbCrLf &
                                "Click CANCEL if you do not want to record the transaction.",
                               "Info", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = System.Windows.Forms.DialogResult.Cancel) Then
                With inventoryAutoBuild
                    .tbxDocNum.Text = ""
                    .tbxDocNum.Select()
                End With
            Else

            End If
        Else

        End If
    End Sub

    Public Sub goodsAutoBuildConfirmationQry(user As String, pass As String)
        SQL.AddParam("@user", user)
        SQL.AddParam("@pas", pass)
        '--NEW ITEM FORM FOR CONFIRMATION. DAPAT MO APPROVED ANG SUPERVISOR ACCOUNT.
        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
	                    INNER JOIN tblEmpDetails T1 ON T1.id = T0.empId
	                    INNER JOIN tblRoles T2 ON T1.role_id = T2.id 
                        WHERE (T0.username = @user) 
                        AND (T0.pass = @pas)
	                    AND (T1.status = 1)
                        AND (T1.role_id IN (20016, 20017,20020));") '--approval procedure approves only by the supervisor or manager

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With goodsAutoBuild
                For Each r As DataRow In SQL.DBDT.Rows
                    .apprvdBy = r(0)
                Next
                newHome.Enabled = True
                .Enabled = True
                goodsAutoBuildConfirmation.Dispose()
                .addDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            With goodsAutoBuildConfirmation
                .tbxUser.Text = ""
                .tbxPass.Text = ""
                .lblError.Visible = False
                .tbxUser.Select()
                .tbxPass.Select() '--added script
            End With
            Exit Sub
        End If
    End Sub

    Public Sub fetchBatchDatasMain(id As String)
        SQL.AddParam("@id", id)

        SQL.ExecQueryDT("SELECT * FROM tblBatches
                        WHERE id = @id;")

        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT > 0 Then
            With updateBatchMain
                For Each r As DataRow In SQL.DBDT.Rows
                    If Not IsDBNull(r(1)) Then
                        .tbxBatchCode.Text = r(1)
                    Else
                        .tbxBatchCode.Text = ""
                    End If

                    If Not IsDBNull(r(2)) Then
                        .tbxDescrip.Text = r(2)
                    Else
                        .tbxDescrip.Text = ""
                    End If
                Next
            End With
        Else
            Exit Sub
        End If
    End Sub

    Public Sub updateBatchMainQry(id As String, code As String, rmks As String, area As String, encBy As String, apprv As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@code", code)

        SQL.ExecQueryDT("SELECT * FROM tblBatches
                        WHERE id <> @id AND batchName = @code;")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Successfully updated.")
        If SQL.RecordCountDT > 0 Then
            With updateBatchMain
                .tbxBatchCode.Text = ""
                .tbxBatchCode.Select()
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .Enabled = True
                .Focus()
                .Select()
            End With
        Else
            With updateBatchMain
                .lblErr.Visible = False
            End With
            SQL.AddParam("@id", id)
            SQL.AddParam("@code", code.ToUpper)
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@area", area)
            SQL.AddParam("@encBy", encBy)
            SQL.AddParam("@apprv", apprv)
            SQL.ExecQueryDT("UPDATE tblBatches
                            SET
                            batchName = @code
                            ,remarks = @rmks
                            ,encBy = @encBy
                            ,encDate = GETDATE()
                            ,approvBy = @apprv
                            WHERE id = @id;")
            If SQL.HasException(True) Then Exit Sub

            updateBatchMain.Dispose()
            newHome.Enabled = True
            With viewBatch
                .Enabled = True
                .Focus()
                .Select()
                .loadLV()
                .lvBatches.SelectedItems.Clear()
                .btnDetails.Enabled = False
            End With
        End If
    End Sub

    Public Sub fethcUpdateLocMainDatas(id As String, area As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE id = @id;")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With updateLocationMain
                    .tbxLocName.Text = r(1)
                    .tbxDescrip.Text = r(2)
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub updateLocMainQry(id As String, name As String, rmks As String, encby As String, area As String, dateModi As String, apprv As String)
        SQL.AddParam("@id", id)
        SQL.AddParam("@name", name)
        SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE id <> @id and [name] = @name and areaId = @area;")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Successfully updated.")
        If SQL.RecordCountDT > 0 Then
            With updateLocationMain
                .tbxLocName.Text = ""
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .Enabled = True
                .Focus()
                .Select()
            End With
        Else
            With updateLocationMain
                .lblErr.Visible = False
            End With
            SQL.AddParam("@id", id)
            SQL.AddParam("@name", name.ToUpper)
            SQL.AddParam("@rmks", rmks)
            SQL.AddParam("@encBy", encby)
            SQL.AddParam("@date", dateModi)
            SQL.AddParam("@apprv", apprv)
            SQL.ExecQueryDT("UPDATE tblLocations
                                SET
                                [name] = @name
                                ,remarks = @rmks
                                ,encby = @encBy
                                ,lastModified = @date
                                ,apprvBy = @apprv
                              WHERE id = @id;")
            If SQL.HasException(True) Then Exit Sub

            With updateLocationMain
                newHome.Enabled = True
                viewLocations.Enabled = True
                .Dispose()
            End With

            With viewLocations
                .loadLV()
                .Focus()
                .Select()
            End With
        End If
    End Sub

    Public Sub loadModuleTypesCBX(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDS("SELECT DISTINCT(T.[MODULE])
FROM(SELECT
(SELECT remarks FROM tblDocumentTypes WHERE id = T1.docType) AS [DOCUMENT TYPE]
,T1.docNo AS [DOCUMENT NUMBER]
,(SELECT CONCAT(COALESCE([name]+ ',',''),COALESCE(lastname+ '',''),COALESCE(middlename,''))  

FROM tblEmpDetails WHERE id = T1.encBy) AS [ENCODED BY]
,T1.docDate AS [DOCUMENT DATE]
,T1.encDate AS [ENCODED DATE]
,DATEDIFF(DAY,T1.docDate,T1.encDate) AS [DIFFERENCE]
,'Receive' AS [MODULE]
FROM tblRecvHeaders T1
UNION ALL
SELECT
(SELECT remarks FROM tblDocumentTypes WHERE id = T1.docType) AS [DOCUMENT TYPE]
,T1.docNo AS [DOCUMENT NUMBER]
,(SELECT CONCAT(COALESCE([name]+ ',',''),COALESCE(lastname+ '',''),COALESCE(middlename,''))  
FROM tblEmpDetails WHERE id = T1.encBy) AS [ENCODED BY]
,T1.docDate AS [DOCUMENT DATE]
,T1.encDate AS [ENCODED DATE]
,DATEDIFF(DAY,T1.docDate,T1.encDate) AS [DIFFERENCE]
,'Release' AS [MODULE]
FROM tblRelsHeader T1
UNION ALL


SELECT
'Transfer Document' AS [DOCUMENT TYPE]
,T1.transferId AS [DOCUMENT NUMBER]
,(SELECT CONCAT(COALESCE([name]+ ', ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
FROM tblEmpDetails WHERE id = T1.encBy) AS [ENCODED BY]
,T1.dateTransfer AS [DOCUMENT DATE]
,T1.encDate AS [ENCODED DATE]
,DATEDIFF(DAY,T1.dateTransfer,T1.encDate) AS [DIFFERENCE]
,'Internal Whouse Transfer' AS [MODULE]
FROM tblIntraTransferHeaders T1
UNION ALL

SELECT
'Entry Number' AS  [DOCUMENT TYPE]
,T1.id AS [DOCUMENT NUMBER]
,(SELECT CONCAT(COALESCE([name]+ ', ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
FROM tblEmpDetails WHERE id = T1.encBy) AS [ENCODED BY]
,T1.transDate AS [DOCUMENT DATE]
,T1.encDate AS [ENCODED DATE]
,DATEDIFF(DAY,T1.transDate,T1.encDate) AS [DIFFERENCE]
,'New Count' AS [MODULE]
FROM tblInventoryAdjustHeaders T1

UNION ALL


SELECT
'Reference Number' AS [DOCUMENT TYPE]
,T1.docNum AS [DOCUMENT NUMBER]
,(SELECT CONCAT(COALESCE([name]+ ', ',''),COALESCE(lastname+ ' ',''),COALESCE(middlename,''))  
FROM tblEmpDetails WHERE id = T1.encBy) AS [ENCODED BY]
,T1.transDate AS [DOCUMENT DATE]
,T1.encDate AS [ENCODED DATE]
,DATEDIFF(DAY,T1.transDate,T1.encDate) AS [DIFFERENCE]
,'New Inventory' AS [MODULE]
FROM tblProductsABHeaders T1

) AS T;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub


    Public Sub loadDateAnalysisArea(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("SELECT area FROM tblareas T0 WHERE T0.isInternal = 1
                        ORDER BY id ASC;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub
        End If
    End Sub


    Public Sub loadConsolidateInventory(lv As ListView, number As String, descrip As TextBox, area As String)
        SQL.AddParam("@id", number)
        SQL.AddParam("@des", "%" & descrip.Text & "%")
        SQL.AddParam("@area", area)
        SQL.ExecQueryDT("SELECT TOP 100 '0' as dummy,T0.goodId,T1.goodDes,T2.area,SUM(T0.qty)
						FROM tblProducts T0
						INNER JOIN tblGoods T1 ON T0.goodId=T1.id
						INNER JOIN tblAreas T2 ON T0.areaId =T2.id
						WHERE (T0.goodId like '%'+ @id +'%' or @id='')
						AND (t1.goodDes LIKE @des or @des ='')
						AND (t2.area = @area or @area='')
						GROUP BY T0.goodId,T0.areaId,T1.goodDes,T2.area
						HAVING SUM(T0.qty)>0
						ORDER BY T0.goodId ASC;")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                End With
            Next
        End If
    End Sub

    'Public Sub loadAllConsolidateInventory(lv As ListView, number As TextBox, descrip As TextBox, area As String)
    '    SQL.DBCon.Close()
    '    SQL.AddParam("@id", number.Text & "%")
    '    SQL.AddParam("@des", "%" & descrip.Text & "%")
    '    SQL.AddParam("@area", area & "%")
    '    lv.Items.Clear()
    '    SQL.ExecQueryDT("SELECT 									
    '                    T0.goodId									
    '                    ,(SELECT goodDes FROM tblGoods WHERE id = T0.goodId)									
    '                    ,(SELECT area FROM tblAreas WHERE id = T0.areaId)									
    '                    ,SUM(T0.qty)									
    '                    FROM tblProducts T0									
    '                    WHERE									
    '                    T0.goodId LIKE @id									
    '                    AND (SELECT goodDes FROM tblGoods WHERE id = T0.goodId) LIKE @des									
    '                    AND (SELECT area FROM tblAreas WHERE id = T0.areaId) LIKE @area									
    '                    GROUP BY T0.goodId									
    '                    ,T0.areaId									
    '                    HAVING SUM(T0.qty) > 0									
    '                    ORDER BY (SELECT goodId FROM tblGoods WHERE id = T0.goodId) ASC;")

    '    If SQL.HasException(True) Then Exit Sub
    '    lv.Items.Clear()
    '    lv.View = View.Details
    '    lv.FullRowSelect = True
    '    lv.GridLines = True
    '    Dim itemNew As New ListViewItem
    '    If SQL.RecordCountDT <> 0 Then
    '        For Each r As DataRow In SQL.DBDT.Rows
    '            itemNew = lv.Items.Add(r(0))
    '            With itemNew
    '                .SubItems.Add(r(1))
    '                .SubItems.Add(r(2))
    '                .SubItems.Add(r(3))
    '            End With
    '        Next
    '    End If

    'End Sub




    Public Sub loadCBXforConsolidatedInv(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDS("SELECT DISTINCT  T0.id,  T0.area
                        FROM tblAreas T0
                        INNER JOIN tblProducts T1 ON T0.id=T1.areaId
                        ORDER BY T0.id ASC
                        ;")
        'SQL.ExecQueryDS("SELECT 
        '                DISTINCT(SELECT area FROM tblAreas WHERE id = T0.areaId)
        '                FROM tblProducts T0;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(1))
            Next
            cbx.SelectedIndex = -1
        Else

        End If
    End Sub

    Public Sub loadLVforConsoInvDetails(lv As ListView, prodId As String, area As String, prodBatch As String, prodLoc As String)
        SQL.AddParam("@id", prodId)
        SQL.AddParam("@area", area)
        SQL.AddParam("@b", prodBatch & "%")
        SQL.AddParam("@l", prodLoc & "%")

        SQL.ExecQueryDT("SELECT 'dummy' as dummy 
                        ,T0.goodId
                        ,(SELECT goodDes FROM tblGoods WHERE id = T0.goodId)
                        ,(SELECT batchName FROM tblBatches WHERE id = T0.batchId)
                        ,(SELECT [name] FROM tblLocations WHERE id = T0.locationId)
                        ,T0.qty
                        ,T0.qty - 
                        ISNULL((SELECT 1 * SUM(qty) FROM tblReserveReleaseDetails x1 
                        JOIN tblReserveReleaseHeaders x2 ON x2.id = x1.transId
                        WHERE prodId = T0.id AND x2.[status] = 'P'),0) AS [Available]
                        ,T0.Expirationdate
                        FROM tblProducts T0 
                        WHERE goodId like '%'+ @id +'%'
          
                        AND (SELECT area FROM tblAreas WHERE id = T0.areaId) = @area
                        AND (SELECT batchName FROM tblBatches WHERE id = T0.batchId) LIKE @b
                        AND (SELECT [name] FROM tblLocations WHERE id = T0.locationId) LIKE @l;")
        If SQL.HasException(True) Then Exit Sub

        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    '.SubItems.Add(r(5))
                    If r.IsNull(6) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(6))
                    End If

                    If r.IsNull(7) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(7))
                    End If
                End With
            Next
        End If

        SQL.AddParam("@id", prodId)
        SQL.AddParam("@area", area)
        SQL.AddParam("@b", prodBatch & "%")
        SQL.AddParam("@l", prodLoc & "%")
        SQL.ExecQueryDT("SELECT 
                        SUM(T0.qty)
                        FROM tblProducts T0 
                        WHERE T0.goodId = @id
                        AND (SELECT area FROM tblAreas WHERE id = T0.areaId) = @area
                        AND (SELECT batchName FROM tblBatches WHERE id = T0.batchId) LIKE @b
                        AND (SELECT [name] FROM tblLocations WHERE id = T0.locationId) LIKE @l;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With consolidatedInvDetails
                    If IsDBNull(r(0)) Then
                        .tbxTotalQty.Text = "0"
                    Else
                        .tbxTotalQty.Text = r(0)
                    End If
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub validateUpdateBatchMain(user As String, pass As String)
        SQL.AddParam("@user", user)
        SQL.AddParam("@pas", pass)

        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
	                    INNER JOIN tblEmpDetails T1 ON T1.id = T0.empId
	                    INNER JOIN tblRoles T2 ON T1.role_id = T2.id
                        WHERE T0.username = @user AND T0.pass = @pas
	                    AND T1.status = 1
	                    AND (T1.role_id IN (20016, 20017,20020));") '--approval procedure approves only by the supervisor or manager
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With updateBatchMain
                For Each r As DataRow In SQL.DBDT.Rows
                    .apprvdBy = r(0)
                Next
                updateBatchConfirmation.Dispose()
                .addDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
            With updateBatchConfirmation
                .tbxUsername.Text = ""
                .tbxPass.Text = ""
                .lblErr.Visible = False
                .tbxUsername.Select()
            End With
            Exit Sub
        End If
    End Sub

    Public Sub validateUpdateLocationMain(user As String, pass As String)
        SQL.AddParam("@user", user)
        SQL.AddParam("@pas", pass)
        SQL.ExecQueryDT("SELECT DISTINCT T0.empId FROM tblUsers T0
	                    INNER JOIN tblEmpDetails T1 ON T1.id = T0.empId
                        INNER JOIN tblRoles T2 ON T1.role_id = T2.id
	                    WHERE T0.username = @user 
                        AND T0.pass = @pas
	                    AND T1.status = 1
	                    AND (T1.role_id IN (20016, 20018,20020));")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With updateLocationMain
                For Each r As DataRow In SQL.DBDT.Rows
                    .apprvBy = r(0)
                Next
                updateLocationConfrimation.Dispose()
                .updateDatas()
            End With
        Else
            MessageBox.Show("Not valid account.", "Error",
                           MessageBoxButtons.OK, MessageBoxIcon.Error)
            With updateLocationConfrimation
                .tbxUser.Text = ""
                .tbxPass.Text = ""
                .lblErr.Visible = False
                .tbxUser.Select()
            End With
            Exit Sub
        End If
    End Sub

    Public Sub loadDisplayDocuTypesLV(lv As ListView, docName As String, rmks As String, isIn As String, isPrim As String)
        SQL.AddParam("@docName", docName)
        SQL.AddParam("@rmks", rmks)
        SQL.AddParam("@isIn", isIn)
        SQL.AddParam("@isPrim", isPrim)

        SQL.ExecQueryDT("SELECT
                        T0.id
						,T0.docName
                        ,T0.remarks
                        --,T0.[date]
                        ,T0.isIn
                        ,T0.isPrimary
                        FROM tblDocumentTypes T0
                        WHERE T0.isIn LIKE @isIn
                        AND T0.remarks LIKE @rmks
                        AND T0.isPrimary LIKE @isPrim
                        AND T0.docName LIKE @docName ORDER BY T0.docName ASC;")

        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    If r(3) = "1" Then
                        .SubItems.Add("Receiving")
                    Else
                        .SubItems.Add("Releasing")
                    End If
                    If r(4) = "1" Then
                        .SubItems.Add("Primary")
                    Else
                        .SubItems.Add("Secondary")
                    End If
                End With
            Next
        End If
    End Sub

    Public Sub generateIGLRSum(ByVal prodDes As String _
                                , ByVal str As Date _
                                , ByVal stp As Date _
                                , ByVal areaName As String _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = str
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)



        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("des")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        ' Clear any existing parameter values
        crParameterValues.Clear()

        ' Split the pdnumn string into multiple values
        Dim pdnumnValues As String() = prodDes.Split(","c) ' Assuming values are comma-separated

        For Each value As String In pdnumnValues
            crParameterDiscreteValue = New ParameterDiscreteValue()
            crParameterDiscreteValue.Value = value.Trim() ' Trim any leading/trailing spaces
            crParameterValues.Add(crParameterDiscreteValue)
        Next

        ' Apply the updated parameter values
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterDiscreteValue.Value = areaName
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub

    Public Sub generateIGLR(ByVal pdnumn As String _
                                , ByVal str As Date _
                                , ByVal stp As Date _
                                , ByRef areaName As String _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = str
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        'crParameterDiscreteValue.Value = pdnumn
        'crParameterFieldDefinitions =
        'cryRpt.DataDefinition.ParameterFields
        'crParameterFieldDefinition =
        'crParameterFieldDefinitions.Item("pdnumber")
        'crParameterValues = crParameterFieldDefinition.CurrentValues

        'crParameterValues.Clear()
        'crParameterValues.Add(crParameterDiscreteValue)
        'crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("pdnumber")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        ' Clear any existing parameter values
        crParameterValues.Clear()

        ' Split the pdnumn string into multiple values
        Dim pdnumnValues As String() = pdnumn.Split(","c) ' Assuming values are comma-separated

        For Each value As String In pdnumnValues
            crParameterDiscreteValue = New ParameterDiscreteValue()
            crParameterDiscreteValue.Value = value.Trim() ' Trim any leading/trailing spaces
            crParameterValues.Add(crParameterDiscreteValue)
        Next

        ' Apply the updated parameter values
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterDiscreteValue.Value = areaName
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub


    Public Sub generateIGLRall(ByVal pdnumn As String _
                                , ByVal str As Date _
                                , ByVal stp As Date _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = str
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("PD")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        ' Clear any existing parameter values
        crParameterValues.Clear()

        ' Split the pdnumn string into multiple values
        Dim pdnumnValues As String() = pdnumn.Split(","c) ' Assuming values are comma-separated

        For Each value As String In pdnumnValues
            crParameterDiscreteValue = New ParameterDiscreteValue()
            crParameterDiscreteValue.Value = value.Trim() ' Trim any leading/trailing spaces
            crParameterValues.Add(crParameterDiscreteValue)
        Next

        ' Apply the updated parameter values
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub
    Public Sub generateDateAnalysis(ByVal str As Date _
                                    , ByVal stp As Date _
                                    , ByVal type As String _
                                    , ByVal areaName As String _
                                    , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("strDate")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = str
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("stpDate")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = stp
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("type")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = type
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = areaName
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub


    Public Sub generateDateallAnalysis(ByVal str As Date _
                                    , ByVal stp As Date _
                                    , ByVal type As String _
                                    , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("strDate")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = str
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("stpDate")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = stp
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("type")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterDiscreteValue.Value = type
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub
    Public Sub loadAreasInReport(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("SELECT area FROM tblareas T0 WHERE T0.isInternal = 1
                        ORDER BY id ASC;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub


    Public Sub loadAreasInReportr(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("SELECT area FROM tblareas T0 WHERE T0.isInternal = 1
                        ORDER BY id ASC;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub
    Public Sub loadAreasInReportdetailed(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("SELECT area FROM tblareas T0 WHERE T0.isInternal = 1
                        ORDER BY id ASC;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub


        End If
    End Sub
    Public Sub loadAreasInReportIGLR(cbx As ComboBox)
        cbx.Items.Clear()
        SQL.ExecQueryDT("SELECT area FROM tblareas T0 WHERE T0.isInternal = 1
                        ORDER BY id ASC;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub
        End If
    End Sub

    Public Sub addABProd_BatchForm(code As String, des As String, encBy As String)
        SQL.AddParam("@code", code)
        SQL.ExecQueryDT("SELECT * FROM tblBatches
                        WHERE batchName = @code;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With autoBuildSubFormBatch
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxBatchName.Text = ""
                .tbxBatchName.Select()
            End With
            Exit Sub
        Else
            SQL.AddParam("@code", code.ToUpper)
            SQL.AddParam("@des", des)
            SQL.AddParam("@enc", encBy)

            SQL.ExecQueryDT("INSERT INTO tblBatches (batchName,remarks,encBy,encDate)
                            VALUES (@code, @des,@enc, GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Batch code successfully added.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With autoBuildSubFormBatch
                .Close()
            End With
        End If
    End Sub

    Public Sub addABProd_LocationForm(code As String, des As String, enc As String, area As String)
        SQL.AddParam("@code", code)
        SQL.AddParam("@area", area)

        SQL.ExecQueryDT("SELECT * FROM tblLocations
                        WHERE [name] = @code AND areaId = @area;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With autoBuildSubFormLocation
                .tbxLocName.Text = ""
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
                .tbxLocName.Select()
            End With
            Exit Sub
        Else
            SQL.AddParam("@code", code.ToUpper)
            SQL.AddParam("@des", des)
            SQL.AddParam("@enc", enc)
            SQL.AddParam("@area", area)

            SQL.ExecQueryDT("INSERT INTO tblLocations ([name], remarks,encby,areaId,lastModified)
                            VALUES (@code, @des, @enc,@area,GETDATE());")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Location name successfully added.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            With autoBuildSubFormLocation
                .Close()
            End With
        End If
    End Sub

    Public Sub fetchProductIDandDesIGLRSpecific(tbxProdId As String, tbxProdDes As String)
        If Not String.IsNullOrWhiteSpace(tbxProdId) Then
            SQL.AddParam("@id", tbxProdId)
            SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE id = @id ")
            If SQL.HasException(True) Then Exit Sub
            If SQL.RecordCountDT > 0 Then
                With IGLRSpecific
                    For Each r As DataRow In SQL.DBDT.Rows
                        .tbxProdNum.Text = ""
                        .tbxProdDes.Text = ""
                        .tbxProdNum.Text = r(0)
                        .tbxProdDes.Text = r(1)
                    Next
                End With
            Else
                With IGLRSpecific
                    .tbxProdDes.Text = ""
                    .tbxProdNum.Text = ""
                    Exit Sub
                End With
            End If
        ElseIf Not String.IsNullOrWhiteSpace(tbxProdDes) Then
            SQL.AddParam("@des", tbxProdDes)
            SQL.ExecQueryDT("SELECT * FROM tblGoods WHERE goodDes = @des;")
            If SQL.HasException(True) Then Exit Sub
            If SQL.RecordCountDT > 0 Then
                With IGLRSpecific
                    For Each r As DataRow In SQL.DBDT.Rows
                        .tbxProdNum.Text = ""
                        .tbxProdDes.Text = ""
                        .tbxProdNum.Text = r(0)
                        .tbxProdDes.Text = r(1)
                    Next
                End With
            Else
                With IGLRSpecific
                    .tbxProdDes.Text = ""
                    .tbxProdNum.Text = ""
                    Exit Sub
                End With
            End If
        End If
    End Sub

    Public Sub FetchDataINCBXIGLRSpecific(tbxprodId As String, cbxArea As ComboBox, str As Date, stp As Date)
        ', cbxBatch As ComboBox, cbxLoc As ComboBox
        'cbx Area
        cbxArea.Items.Clear()
        SQL.AddParam("@id", tbxprodId)
        SQL.AddParam("@str", str)
        SQL.AddParam("@stp", stp)
        SQL.ExecQueryDS("sp_FetchProductAREAINGLSpecific @id, @str, @stp;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbxArea.Items.Add(r(0))
            Next
            cbxArea.Enabled = True
            cbxArea.SelectedIndex = 0
        Else
            cbxArea.SelectedIndex = -1
            cbxArea.Enabled = False
            With IGLRSpecific
                .cbxBatch.SelectedIndex = -1
                .cbxLoc.SelectedIndex = -1
                .cbxBatch.Enabled = False
                .cbxLoc.Enabled = False
            End With
            'Exit Sub
        End If


    End Sub

    Public Sub fetchBatchAndLocationINGLSpecificCBX(id As String, area As String, cbxBatch As ComboBox, cbxLoc As ComboBox, str As Date, stp As Date)
        'cbx Batch   
        cbxBatch.Items.Clear()
        SQL.AddParam("@id", id)
        SQL.AddParam("@area", area)
        SQL.AddParam("@str", str)
        SQL.AddParam("@stp", stp)
        SQL.ExecQueryDS("sp_FetchProductBATCHINGLSpecific @id,@area,@str,@stp;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbxBatch.Items.Add(r(0))
            Next
            cbxBatch.Enabled = True
            cbxBatch.SelectedIndex = 0
        Else
            cbxBatch.SelectedIndex = -1
            cbxBatch.Enabled = False
            'Exit Sub
        End If

        'cbx Location   
        cbxLoc.Items.Clear()
        SQL.AddParam("@id", id)
        SQL.AddParam("@area", area)
        SQL.AddParam("@str", str)
        SQL.AddParam("@stp", stp)
        SQL.ExecQueryDS("sp_FetchProductLOCINGLSpecific @id,@area,@str, @stp;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbxLoc.Items.Add(r(0))
            Next
            cbxLoc.Enabled = True
            cbxLoc.SelectedIndex = 0
        Else
            cbxLoc.SelectedIndex = -1
            cbxLoc.Enabled = False
            'Exit Sub
        End If
    End Sub

    Public Sub generateIGLRSpecific(ByVal prodnum As String _
                               , ByVal str As Date _
                               , ByVal stp As Date _
                               , ByVal areaName As String _
                               , ByVal batch As String _
                               , ByVal loc As String _
                               , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = str
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = prodnum
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("id")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterDiscreteValue.Value = areaName
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = batch
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("batch")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = loc
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("loc")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub

    Public Sub addLocation_IntraTransfer(x As String, y As String, areaId As String, userId As String)
        SQL.AddParam("@name", x)
        SQL.AddParam("@area", areaId)

        SQL.ExecQueryDT("SELECT * FROM tblLocations WHERE [name] = @name AND areaId = @area;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            With addLocationIntraTransfer
                .tbxLocName.Text = ""
                .lblErr.Visible = True
                .lblErr.Text = "Data already exists."
            End With
        Else
            SQL.AddParam("@name", x.ToUpper)
            SQL.AddParam("@des", y)
            SQL.AddParam("@area", areaId)
            SQL.AddParam("@id", userId)
            SQL.ExecQueryDT("INSERT INTO tblLocations
                            ([name], remarks, encby, areaId, lastModified)
                            VALUES
                            (@name, @des, @id, @area, GETDATE());")
            If SQL.HasException(True) Then Exit Sub

            With addLocationIntraTransfer
                .Close()
            End With
        End If
    End Sub

    Public Sub fetchCBXLocationInLedgerReportSpecific(id As String, area As String, batch As String, strDate As Date, stpDate As Date, cbx As ComboBox)
        cbx.Items.Clear()

        SQL.AddParam("@id", id)
        SQL.AddParam("@area", area)
        SQL.AddParam("@batch", batch)
        SQL.AddParam("@str", strDate)
        SQL.AddParam("@stp", stpDate)

        SQL.ExecQueryDS("[dbo].[sp_FetchLOCATION] @id, @area, @str, @stp, @batch;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next

        Else
            Exit Sub
        End If
        cbx.SelectedIndex = 0
    End Sub

    Public Sub fetchCBXBatchInLedgerReportSpecific(id As String, area As String, loc As String, strDate As String, stpDate As String, cbx As ComboBox)
        cbx.Items.Clear()

        SQL.AddParam("@id", id)
        SQL.AddParam("@area", area)
        SQL.AddParam("@loc", loc)
        SQL.AddParam("@str", strDate)
        SQL.AddParam("@stp", stpDate)

        SQL.ExecQueryDS("[dbo].[sp_FetchBATCH] @id, @area, @str, @stp,@loc;")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDS > 0 Then
            For Each r As DataRow In SQL.DBDS.Tables(0).Rows
                cbx.Items.Add(r(0))
            Next
            cbx.SelectedIndex = 0
        Else
            Exit Sub
        End If
        cbx.SelectedIndex = 0
    End Sub

    Public Sub loadProductListforItemListSummary(cbx As CheckedListBox, tbx As TextBox, tby As TextBox)
        cbx.Items.Clear()
        SQL.AddParam("@des", tby.Text & "%")
        SQL.AddParam("@id", tbx.Text & "%")
        SQL.ExecQueryDT("SELECT CONCAT('(',id, ') ', goodDes) FROM tblGoods 
                        WHERE id IN 
                        (SELECT DISTINCT(T0.goodId) FROM tblProducts T0
						JOIN tblGoods T1 ON T1.id = T0.goodId AND T0.goodId LIKE @id AND T1.goodDes LIKE @des);")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                cbx.Items.Add(r(0))
            Next
        Else
            Exit Sub
        End If
    End Sub


    Public Sub generateItemListSummary(ByVal prodId As String _
                                , ByVal areaName As String _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        crParameterDiscreteValue.Value = prodId
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("id")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = areaName
        crParameterFieldDefinitions =
        cryRpt.DataDefinition.ParameterFields
        crParameterFieldDefinition =
        crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues

        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub

    Public Sub loadConsolidateInventory2(lv As ListView, number As TextBox, descrip As TextBox, locStartFrom As TextBox, batch As TextBox, area As String)
        SQL.AddParam("@igoodid", number.Text & "%")
        SQL.AddParam("@vgooddes", "%" & descrip.Text & "%")
        SQL.AddParam("@vlocation", "%" & locStartFrom.Text & "%")
        'SQL.AddParam("@locEndTo", "%" & locEndTo.Text & "%")
        SQL.AddParam("@vbatches", "%" & batch.Text & "%")
        SQL.AddParam("@varea", area & "%")
        '        SQL.ExecQueryDT("SELECT TOP 500
        '                        T0.goodId
        '                        ,(SELECT goodDes FROM tblGoods WHERE id = T0.goodId)
        '                        ,(SELECT area FROM tblAreas WHERE id = T0.areaId)
        '                        ,(SELECT batchName FROM tblBatches WHERE id=t0.batchId)
        '                        ,(SELECT name FROM tblLocations WHERE id=t0.locationId)
        '                        ,SUM(T0.qty)
        '                        FROM tblProducts T0
        '                        WHERE T0.goodId LIKE @id
        '                        AND (SELECT goodDes FROM tblGoods WHERE id = T0.goodId) LIKE @des
        '                        AND (SELECT area FROM tblAreas WHERE id = T0.areaId) LIKE @area
        '                        AND (SELECT name FROM tblLocations WHERE id=t0.locationId) LIKE @locStartFrom  
        '                        AND (SELECT batchName FROM tblBatches WHERE id=t0.batchId) LIKE @batch
        '                        GROUP BY T0.goodId
        '                        ,T0.areaId
        '                        ,T0.batchId
        '                        ,T0.locationId
        '                        HAVING SUM(T0.qty) > 0
        'ORDER BY (SELECT goodId FROM tblGoods WHERE id = T0.goodId) ASC")
        SQL.ExecQueryDT("EXEC loadConsolidateInventory2 @igoodid=@igoodid,@vgooddes=@vgooddes,@varea=@varea,@vbatches=@vbatches,@vlocation=@vlocation")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                End With
            Next
        End If




    End Sub

    Public Sub loadConsolidateInventory4(lv As ListView, number As String, descrip As TextBox, locStartFrom As TextBox, batch As TextBox, area As String)
        SQL.AddParam("@igoodid", "%" & number & "%")
        SQL.AddParam("@vgooddes", "%" & descrip.Text & "%")
        SQL.AddParam("@vlocation", locStartFrom.Text)
        'SQL.AddParam("@locEndTo", "%" & locEndTo.Text & "%")
        SQL.AddParam("@vbatches", "%" & batch.Text & "%")
        SQL.AddParam("@varea", area)

        SQL.ExecQueryDT("EXEC loadConsolidateInventory4 @igoodid=@igoodid,@vgooddes=@vgooddes,@varea=@varea,@vbatches=@vbatches,@vlocation=@vlocation")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                End With
            Next
        End If

    End Sub

    Public Sub loadcountdata(ls As ListView, t1 As String, t2 As String, t3 As String, t4 As String, colSelected As Integer)
        Dim CURRENTAREA As Integer = newHome.areaId

        ls.Items.Clear()
        If newHome.roleId = "20020" Then

            SQL.AddParam("@id", t1)
            SQL.AddParam("@des", t2)
            SQL.AddParam("@loc", t3)
            SQL.AddParam("@batch", t4)
            SQL.ExecQueryDT("SELECT TOP 500 
            t0.id
            ,t0.goodId as [PRODUCT ID]
            ,t1.goodDes as [DESCRIPTION]
            ,t2.batchName as [BATCH]
            ,t3.name as [LOCATION]
			,T4.area AS [AREA]
            ,t0.Expirationdate as [EXPIRATION DATE]
            ,1 * t0.qty as [ON HAND]
            ,1 * t0.qty as [COUNTED]
            ,1 * (t0.qty - t0.qty) as [DIFFERENCE]
            FROM tblProducts t0
            INNER JOIN tblGoods t1 ON t0.goodId = t1.id
            INNER JOIN tblBatches t2 ON t0.batchId = t2.id
            INNER JOIN tblLocations t3 ON t0.locationId = t3.id
			INNER JOIN tblAreas T4 ON T0.areaId=T4.id
            where (t0.goodId = @id or @id='') AND
            (t1.goodDes LIKE @des or @des='') AND 
            (t2.batchName = @batch or @batch='') AND 
            (t3.name LIKE @loc or @loc='' )")

            If SQL.HasException(True) Then Exit Sub


            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    Dim item As New ListViewItem()
                    Dim idValue As String = r("ID").ToString()
                    item.Text = idValue ' Set the value of the ListViewItem itself to the ID value
                    item.SubItems.Add(r("PRODUCT ID").ToString())
                    item.SubItems.Add(r("DESCRIPTION").ToString())
                    item.SubItems.Add(r("BATCH").ToString())
                    item.SubItems.Add(r("LOCATION").ToString())
                    item.SubItems.Add(r("AREA").ToString())
                    item.SubItems.Add(r("EXPIRATION DATE").ToString())
                    item.SubItems.Add(r("ON HAND").ToString())
                    item.SubItems.Add(r("COUNTED").ToString())
                    item.SubItems.Add(r("DIFFERENCE").ToString())
                    ls.Items.Add(item)
                Next
            Else
                Exit Sub

            End If

        Else


            SQL.AddParam("@id", t1)
            SQL.AddParam("@des", t2)
            SQL.AddParam("@loc", t3)
            SQL.AddParam("@batch", t4)
            SQL.AddParam("@CURRENTAREA", CURRENTAREA)

            SQL.ExecQueryDT("SELECT TOP 500 
            t0.id
            ,t0.goodId as [PRODUCT ID]
            ,t1.goodDes as [DESCRIPTION]
            ,t2.batchName as [BATCH]
            ,t3.name as [LOCATION]
			,T4.area AS [AREA]
            ,t0.Expirationdate as [EXPIRATION DATE]
            ,1 * t0.qty as [ON HAND]
            ,1 * t0.qty as [COUNTED]
            ,1 * (t0.qty - t0.qty) as [DIFFERENCE]
            FROM tblProducts t0
            INNER JOIN tblGoods t1 ON t0.goodId = t1.id
            INNER JOIN tblBatches t2 ON t0.batchId = t2.id
            INNER JOIN tblLocations t3 ON t0.locationId = t3.id
			INNER JOIN tblAreas T4 ON T0.areaId=T4.id
            where (t0.goodId = @id or @id='') AND
            (t1.goodDes LIKE @des or @des='') AND 
            (t2.batchName = @batch or @batch='') AND 
            (t3.name LIKE @loc or @loc='' ) And
            (T4.id=@CURRENTAREA)")

            If SQL.HasException(True) Then Exit Sub


            If SQL.RecordCountDT > 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    Dim item As New ListViewItem()
                    Dim idValue As String = r("ID").ToString()
                    item.Text = idValue ' Set the value of the ListViewItem itself to the ID value
                    item.SubItems.Add(r("PRODUCT ID").ToString())
                    item.SubItems.Add(r("DESCRIPTION").ToString())
                    item.SubItems.Add(r("BATCH").ToString())
                    item.SubItems.Add(r("LOCATION").ToString())
                    item.SubItems.Add(r("AREA").ToString())
                    item.SubItems.Add(r("EXPIRATION DATE").ToString())
                    item.SubItems.Add(r("ON HAND").ToString())
                    item.SubItems.Add(r("COUNTED").ToString())
                    item.SubItems.Add(r("DIFFERENCE").ToString())
                    ls.Items.Add(item)
                Next
            Else
                Exit Sub

            End If

        End If

    End Sub

    Public Sub updateloadGoodsToEdit(lv As ListView, tbx As String, tbx1 As String)
        SQL.AddParam("@goodid", tbx)
        SQL.AddParam("@des", tbx1 & "%")
        SQL.ExecQueryDT("SELECT TOP 50 * 
                            FROM tblGoods T0
                            WHERE (T0.id = @goodid  OR @goodid='')
                            AND (T0.goodDes LIKE @des OR @des='')")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))

                With itemNew
                    .SubItems.Add(r(1))
                End With
            Next

        End If
    End Sub


    Public Sub uodateloadArea(id As String, tbx1 As TextBox)

        Try

            SQL.AddParam("@goodid", id)
            SQL.AddParam("@des", tbx1.Text)

            SQL.ExecQueryDT("UPDATE tblGoods SET goodDes=@des WHERE id = @goodid ")
            If SQL.HasException(True) Then Exit Sub

        Catch ex As Exception

        End Try
    End Sub 'load list view


    Public Sub updatereasoncode(txt1 As String, txt2 As String, txt3 As String)

        Try

            SQL.AddParam("@id", txt3)
            SQL.AddParam("@code", txt1)
            SQL.AddParam("@des", txt2)

            SQL.ExecQueryDT("UPDATE tblReasoncode SET description=@des,reasoncode=@code WHERE id like '%'+ @id +'%'")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Successfully Updated...")

            With reasoncode
                .view()
                .txt1.Text = ""
                .txt2.Text = ""
                .TextBox2.Text = ""
            End With

        Catch ex As Exception

        End Try
    End Sub 'load list view


    Public Sub generatereceive(ByVal selectedDocno As String _
                                 , ByVal selectedDoctype As String _
                                 , ByVal selectedSource As String _
                                 , ByVal selectedArea As String _
                                 , ByVal str As Date _
                                 , ByVal stp As Date _
                                 , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("docno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedDoctype
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("doctype")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedSource
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("sender")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub


    Public Sub generaterelease(ByVal selectedDocno As String _
                                , ByVal selectedDoctype As String _
                                , ByVal selectedSource As String _
                                , ByVal selectedArea As String _
                                , ByVal str As Date _
                                , ByVal stp As Date _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("docno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedDoctype
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("doctype")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedSource
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("sender")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtfrom")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtto")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub


    Public Sub generateinventorybuildreg(ByVal selectedDocno As String _
                               , ByVal selectedDoctype As String _
                               , ByVal selectedArea As String _
                               , ByVal str As Date _
                               , ByVal stp As Date _
                               , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("docno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedDoctype
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("entryno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)



        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtfrom")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtto")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub


    Public Sub generateItemRegister(ByVal selectedDocno As String _
                              , ByVal str As Date _
                              , ByVal stp As Date _
                              , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("docno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtfrom")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtto")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub


    Public Sub generateTransferRegister(ByVal selectedDocno As String _
                             , ByVal selectedDoctype As String _
                             , ByVal selectedArea As String _
                             , ByVal str As Date _
                             , ByVal stp As Date _
                             , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("docno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedDoctype
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("entryid")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtfrom")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("dtto")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub



    Public Sub generateCountRegister(ByVal selectedDocno As String _
                           , ByVal selectedArea As String _
                           , ByVal str As Date _
                           , ByVal stp As Date _
                           , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("entrynumber")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        ' Set str parameter value
        crParameterDiscreteValue.Value = str
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("str")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        ' Set stp parameter value
        crParameterDiscreteValue.Value = stp
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("stp")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub



    Public Sub generateProductRegister(ByVal selectedDocno As String _
                          , ByVal selecteddes As String _
                          , ByVal selectedArea As String _
                          , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        'Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("pdno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selecteddes
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("Des")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


    End Sub


    Public Sub generateInventorylistbylocation(ByVal selectedDocno As String _
                         , ByVal selecteddes As String _
                         , ByVal selectedArea As String _
                         , ByVal selectedloc As String _
                         , ByVal selectedbatch As String _
                         , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        'Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("pdno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selecteddes
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("Des")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterDiscreteValue.Value = selectedloc
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("loc")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


        crParameterDiscreteValue.Value = selectedbatch
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("batch")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


    End Sub


    Public Sub generateInventorylist(ByVal selectedDocno As String _
                         , ByVal selecteddes As String _
                         , ByVal selectedArea As String _
                         , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        'Set selectedDocno parameter value
        crParameterDiscreteValue.Value = selectedDocno
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("pdno")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selecteddes
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("Des")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = selectedArea
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub


    Public Sub generateExpiration(ByVal proDuctNo As String _
                                , ByVal des As String _
                                , ByVal date1 As String _
                                , ByVal date2 As String _
                                , ByVal area As String _
                                , ByVal filter As String _
                                , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = proDuctNo
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("PD")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = des
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("des")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = date1
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("datediff")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = date2
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("datediff2")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = area
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("area")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = filter
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("filter")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)
    End Sub


    Public Sub EditExDate()
        Dim pd As Integer

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                pd = r(0)
            Next

            With EditExpirationdate.dgexdate

                For i As Integer = 0 To .Rows.Count - 1 Step +1
                    If .Rows(i).Cells(7).Value <> 0 Then

                        SQL.AddParam("@transId", pd)
                        SQL.AddParam("@pId", .Rows(i).Cells(0).Value.ToString)
                        'SQL.AddParam("@counted", .Rows(i).Cells(6).Value.ToString)
                        'SQL.AddParam("@onhand", .Rows(i).Cells(7).Value.ToString)
                        SQL.AddParam("@status", .Rows(i).Cells(7).Value.ToString)
                        SQL.AddParam("@xd", .Rows(i).Cells(6).Value.ToString)

                        SQL.ExecQueryDT("UPDATE tblProducts
                         SET Expirationdate = @xd
                         WHERE id = @pId;")

                        If SQL.HasException(True) Then Exit Sub

                    End If

                Next
                MessageBox.Show("Successfully Updated")
            End With
        Else
            Exit Sub
        End If

        With EditExpirationdate.dgexdate
            .Refresh()


        End With
    End Sub



    Public Sub addreasoncode(txt1 As String, txt2 As String, txt3 As String)
        SQL.AddParam("@code", txt1)
        SQL.AddParam("@des", txt2)
        SQL.AddParam("@module", txt3)

        SQL.ExecQueryDT("select * from tblReasoncode where reasoncode = @code;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            MessageBox.Show("Reason Code Already Exist...")
        Else
            SQL.AddParam("@code", txt1.ToUpper())
            SQL.AddParam("@des", txt2)
            SQL.AddParam("@module", txt3)
            SQL.ExecQueryDT("INSERT INTO tblReasoncode
                            (reasoncode,description,module)
                            VALUES
                            (@code,@des,@module);")
            If SQL.HasException(True) Then Exit Sub
            MessageBox.Show("Reason Code successfully added.", "Info!", MessageBoxButtons.OK, MessageBoxIcon.Information)

            With reasoncode
                .view()
                .txt1.Text = ""
                .txt2.Text = ""
                .TextBox2.Text = ""
            End With



        End If
    End Sub

    Public Sub viewreasoncode(lv As ListView, txt1 As String, txt2 As String, txt3 As String)
        SQL.AddParam("@code", txt1)
        SQL.AddParam("@des", txt2)
        SQL.AddParam("@mod", txt3)
        SQL.ExecQueryDT("select t1.id,t1.reasoncode,t1.description,t1.module
                        from tblReasoncode t1
                        where (t1.reasoncode LIKE @code or @code = '')
                        and (t1.description LIKE @des or @des='')
                        and (t1.module= @mod or @mod = '')
                        ")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            lv.Items.Clear()
            lv.View = View.Details
            lv.FullRowSelect = True
            lv.GridLines = True
            Dim itemNew As New ListViewItem
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    itemNew = lv.Items.Add(r(0))
                    With itemNew
                        .SubItems.Add(r(1))
                        .SubItems.Add(r(2))
                        .SubItems.Add(r(3))
                    End With
                Next


            End If
        Else

        End If

    End Sub

    Public Sub loadorderpoint(lv As ListView, txt1 As String, txt2 As String)
        SQL.AddParam("@id", txt1)
        SQL.AddParam("@des", "%" & txt2 & "%")

        SQL.ExecQueryDT("SELECT TOP 50 '' as dummy,T0.[Product Number],T0.Description
                            , [dbo].[SVreorderpoit](CEILING(SUM((T0.DAILYSALES*T0.LeadTime)+(T0.LeadTimeAllowance*T0.DAILYSALES))),T0.[Annual Total Sales]) AS REORDER
                            ,t0.ONHAND
                            ,[dbo].[SVremark]( [dbo].[SVreorderpoit](CEILING(SUM((T0.DAILYSALES*T0.LeadTime)+(T0.LeadTimeAllowance*T0.DAILYSALES))),T0.[Annual Total Sales]),T0.ONHAND) as REMARKS
                            ,[dbo].[SVtobuy](T0.ONHAND,T0.[Annual Total Sales]) as TOBUY
                            ,t0.[Annual Total Sales]
                            FROM (
                            SELECT distinct T0.[Product Number],T0.Description,T0.[Annual Total Sales],CEILING(T0.[Annual Total Sales]/312) AS DAILYSALES,SUM(T1.qty)AS ONHAND,T0.LeadTime,T0.LeadTimeAllowance
                            FROM [dbo].[opview] T0
                            INNER JOIN tblProducts T1 ON T0.[Product Number]=T1.goodId
                            GROUP BY T0.[Product Number],T0.Description,T0.[Annual Total Sales],T0.LeadTime,T0.LeadTimeAllowance
                            ) AS T0
                            WHERE (T0.[Product Number]=@id or @id='')
                            AND (T0.Description LIKE @des or @des='')
                            GROUP BY T0.[Product Number],T0.Description,T0.ONHAND,t0.[Annual Total Sales],t0.DAILYSALES")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                    .SubItems.Add(r(7))
                End With
            Next
        End If
    End Sub




    Public Sub loadreordersetup(t1 As String, t2 As String, Optional query As String = "")

        If query = "" Then
            SQL.AddParam("@gid", t1)
            SQL.AddParam("@des", "%" & t2 & "%")
            SQL.ExecQueryDT("SELECT TOP 50 T0.goodId as [Product Number],T0.goodDes as [Description], SUM(T0.CASHTOTAL - T0.CM) as [Annual Total Sales],T1.LeadTime,T1.LeadTimeAllowance,'' as dummy
                                FROM (
                                    SELECT DISTINCT t0.id, T0.goodId,t.goodDes,T0.qty AS [CASHTOTAL],  '' AS CM
                                    FROM tblCash T0
	                                inner join tblGoods t on T0.goodId=t.id
                                    UNION

                                    SELECT DISTINCT t1.id ,T1.goodId,t.goodDes, T1.qty AS [CASHTOTAL], '' AS CM
                                    FROM tblChargeInvoice T1
		                                inner join tblGoods t on T1.goodId=t.id
                                    UNION 
                                    SELECT DISTINCT t2.id,t2.goodid,t.goodDes,'' AS [CASHTOTAL],t2.qty AS CM
                                    FROM tblCm t2
		                                inner join tblGoods t on t2.goodId=t.id
                                ) AS T0
                                INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                                WHERE (T0.goodId like '%'+ @gid +'%' OR @gid='') AND (T0.goodDes LIKE @des OR @des='')
                                GROUP BY T0.goodId, T0.goodDes,T1.LeadTime,T1.LeadTimeAllowance
                                HAVING SUM(T0.CASHTOTAL - T0.CM) >= 0
                                ORDER BY T0.goodId ASC")
            If SQL.HasException(True) Then Exit Sub
        Else
            SQL.ExecQueryDT(query)
        End If


        With reordersetup.dgorder
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT

            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35

            .Columns(0).HeaderText = "PRODUCT NUMBER"
            .Columns(1).HeaderText = "DESCRIPTION"
            .Columns(2).HeaderText = "ANNUAL TOTAL SALES"
            .Columns(3).HeaderText = "LEAD TIME"
            .Columns(4).HeaderText = "LEAD TIME ALLOWANCE"
            'gamit ra para pang validate
            .Columns(5).HeaderText = "STATUS"


            .Columns(0).Width = 150
            .Columns(1).Width = 500
            .Columns(2).Width = 150
            .Columns(3).Width = 150
            .Columns(4).Width = 150
            .Columns(5).Width = 150

            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(4).ReadOnly = True
            '.Columns(5).Visible = False

            .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .CurrentCell = Nothing

        End With

    End Sub


    Public Sub setupOrderpoint()

        With reordersetup.dgorder
            For i As Integer = 0 To .Rows.Count - 2 Step +1
                SQL.AddParam("@gId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@lead", .Rows(i).Cells(4).Value.ToString)
                SQL.AddParam("@leada", .Rows(i).Cells(5).Value.ToString)
                SQL.ExecQueryDT("update tblGoods set LeadTime=@lead,LeadTimeAllowance=@leada where id like '%'+ @gId +'%' ")
                If SQL.HasException(True) Then Exit Sub
            Next
            MessageBox.Show("Successfully Recoreded...")

        End With

        With reordersetup
            .btnAdd.Text = "New Entry"
        End With
    End Sub



    Public Sub loadorderpointsetup(lv As ListView, tbx As String, tbx2 As String)


        SQL.AddParam("@fltr", tbx)
        SQL.AddParam("@fltr2", "%" & tbx2 & "%")
        SQL.ExecQueryDT("SELECT TOP 50 '0' as dummy, T0.goodId as [Product Number],T0.goodDes as [Description], SUM(T0.CASHTOTAL - T0.CM) as [Annual Total Sales],T1.LeadTime,T1.LeadTimeAllowance
                            FROM (
                                SELECT DISTINCT t0.id, T0.goodId,t.goodDes,T0.qty AS [CASHTOTAL],  '' AS CM
                                FROM tblCash T0
	                            inner join tblGoods t on T0.goodId=t.id
                                UNION

                                SELECT DISTINCT t1.id ,T1.goodId,t.goodDes, T1.qty AS [CASHTOTAL], '' AS CM
                                FROM tblChargeInvoice T1
		                            inner join tblGoods t on T1.goodId=t.id
                                UNION 
                                SELECT DISTINCT t2.id,t2.goodid,t.goodDes,'' AS [CASHTOTAL],t2.qty AS CM
                                FROM tblCm t2
		                            inner join tblGoods t on t2.goodId=t.id
                            ) AS T0
                            INNER JOIN tblGoods T1 ON T0.goodId=T1.id
                            WHERE (T0.goodId like '%'+ @fltr +'%' OR @fltr='') AND (T0.goodDes LIKE  @fltr2 OR @fltr2 ='')
                            GROUP BY T0.goodId, T0.goodDes,T1.LeadTime,T1.LeadTimeAllowance
                            HAVING SUM(T0.CASHTOTAL - T0.CM) >= 0
                            ORDER BY T0.goodId ASC
                            ")
        If SQL.HasException(True) Then Exit Sub
        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True

        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))

                End With
            Next
        End If

    End Sub


    Public Sub fetchorderpointdetails(lv As ListView, prodId As String)
        SQL.AddParam("@id", prodId)

        SQL.ExecQueryDT("Select '' as dummy, t0.goodId,t1.goodDes,t2.area,t3.batchName,t4.name,t0.qty,t0.Expirationdate
                            from tblProducts t0
                            inner join tblGoods t1 on t0.goodId=t1.id
                            inner join tblAreas t2 on t0.areaId=t2.id
                            inner join tblBatches t3 on t0.batchId=t3.id
                            inner join tblLocations t4 on t0.locationId=t4.id
							WHERE T0.goodId LIKE '%'+ @id +'%'")
        If SQL.HasException(True) Then Exit Sub

        lv.Items.Clear()
        lv.View = View.Details
        lv.FullRowSelect = True
        lv.GridLines = True
        Dim itemNew As New ListViewItem
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                itemNew = lv.Items.Add(r(0))
                With itemNew
                    .SubItems.Add(r(1))
                    .SubItems.Add(r(2))
                    .SubItems.Add(r(3))
                    .SubItems.Add(r(4))
                    .SubItems.Add(r(5))
                    .SubItems.Add(r(6))
                    If r.IsNull(7) Then
                        .SubItems.Add("N/A")
                    Else
                        .SubItems.Add(r(7))
                    End If
                End With
            Next
        End If

        SQL.AddParam("@id", prodId)

        SQL.ExecQueryDT("SELECT SUM(qty) FROM tblProducts WHERE goodId LIKE '%'+ @id +'%' @id")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT > 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                With Orderpoitdetails
                    If IsDBNull(r(0)) Then
                        .TextBox1.Text = "0"
                    Else
                        .TextBox1.Text = r(0)
                    End If
                End With
            Next
        Else
            Exit Sub
        End If
    End Sub


    Public Sub generateorderpoint(ByVal pd As String _
                                , ByVal des As String _
                                , ByVal flter1 As String _
                                , ByVal flter2 As String _
                                , ByVal remark As String _
                               , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = pd
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("pd")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = des
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("des")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = flter1
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("flter1")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = flter2
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("flter2")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = remark
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("remarks")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


    End Sub





    Public Sub ednotif()
        If newHome.roleId = 20020 Or 20016 Then

            Dim diff As String = ""
            Dim diff1 As String = "3"
            Dim remarks As String = "BUY"
            SQL.AddParam("@diff", diff)
            SQL.AddParam("@diff2", diff1)

            SQL.ExecQueryDT("SELECT CONCAT(COUNT(T2.area),' Items will expire in 3 months.') AS area_count
                            FROM tblProducts T0
                            INNER JOIN tblGoods T1 ON T0.goodId = T1.id
                            INNER JOIN tblAreas T2 ON T0.areaId = T2.id
                            INNER JOIN tblBatches T3 ON T0.batchId = T3.id
                            INNER JOIN tblLocations T4 ON T0.locationId = T4.id
                            WHERE (T0.Expirationdate IS NOT NULL AND T0.Expirationdate != 'No Expiration Date')
                                  AND (ISDATE(T0.Expirationdate) = 1 AND DATEDIFF(MONTH, GETDATE(), TRY_CONVERT(DATE, T0.Expirationdate)) BETWEEN @diff AND @diff2)")

            If SQL.HasException(True) Then Exit Sub
            If SQL.RecordCountDT <> 0 Then

                With exdate
                    ' Clear existing rows in the DataGridView (optional, based on your requirement)
                    .dgview.Rows.Clear()

                    ' Clear existing columns in the DataGridView (optional, based on your requirement)
                    .dgview.Columns.Clear()

                    ' Add a column to the DataGridView
                    .dgview.Columns.Add("Expiring Items", "NOTIFICATIONS")
                    .dgview.Columns(0).Width = 700

                    For Each r As DataRow In SQL.DBDT.Rows
                        If Not IsDBNull(r(0)) Then
                            ' Add item to the DataGridView
                            .dgview.Rows.Add(r(0).ToString())
                        End If
                    Next


                End With
            End If

            SQL.AddParam("@remark", remarks)
            SQL.ExecQueryDT("exec [dbo].[spNotif2] @remark")
            If SQL.HasException(True) Then Exit Sub
            If SQL.RecordCountDT <> 0 Then


                With exdate

                    ' Loop through the DataTable rows and add items to the DataGridView
                    For Each r As DataRow In SQL.DBDT.Rows
                        If Not IsDBNull(r(0)) Then
                            ' Add item to the DataGridView
                            .dgview.Rows.Add(r(0).ToString())
                        End If
                    Next
                End With
            End If

            SQL.ExecQueryDT("exec [dbo].[agingnotif]")
            If SQL.HasException(True) Then Exit Sub
            If SQL.RecordCountDT <> 0 Then
                With exdate

                    ' Loop through the DataTable rows and add items to the DataGridView
                    For Each r As DataRow In SQL.DBDT.Rows
                        If Not IsDBNull(r(0)) Then
                            ' Add item to the DataGridView
                            .dgview.Rows.Add(r(0).ToString())
                        End If
                    Next
                End With
            End If


        Else
            'nothing'
        End If




    End Sub

    'fetch series for Document Number in receiving : Receiving Report
    Public Sub fetchSeriesRecvAndDisplay()
        Dim area As String = ""

        SQL.AddParam("@areaid", newHome.areaId)
        SQL.ExecQueryDT("SElect AreaCode from tblAreas where id =@areaid")
        If SQL.HasException(True) Then Exit Sub
        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                area = r(0)
            Next
        End If

        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRecvHeaders where  docType = 50014 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer


            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = area + "-" + numericPart.ToString

                With recvGoodsMain
                    .series = incrementedSeries
                End With

            End If
        Else

            Dim defaultSeries As String = area & "-1"

            With recvGoodsMain
                .series = defaultSeries
            End With

        End If
    End Sub

    'fetch series for Document Number in receiving :Goods Receipt Form-Receiving
    Public Sub fetchGRFseries()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRecvHeaders where docType = 50005 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")

                With recvGoodsMain
                    .series = incrementedSeries
                End With

            End If
        Else

            Dim defaultSeries As String = "1-00001"

            With recvGoodsMain
                .series = defaultSeries
            End With

        End If

    End Sub

    'fetch series for Document Number in Releasing : Goods Receipt Form
    Public Sub FetchSeriesinReleasing()

        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where docType = 50004 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")

                With releaseGoods
                    .series1 = incrementedSeries
                End With

            End If

        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub

    'fetch series for Document Number in Releasing : Goods Receipt Form : For Company expense
    Public Sub FetchSeriesinReleasingCE()

        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where docType = 50027 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")

                With releaseGoods
                    .series1 = incrementedSeries
                End With

            End If

        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub

    'fetch series for Document Number in Releasing : Purchase Return
    Public Sub FetchSeriesinReleasing1()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where docType = 50008 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")

                With releaseGoods
                    .series1 = incrementedSeries
                End With

            End If
        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub


    'fetch series for Document Number in Releasing : Delivery Receipt
    Public Sub FetchSeriesinReleasing2()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where docType = 50002 order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")

                With releaseGoods
                    .series1 = incrementedSeries
                End With

            End If
        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub
    Public Sub addrecvseries(series As String)
        SQL.AddParam("@series", series)

        SQL.ExecQueryDT("INSERT INTO tblrecvseries
                            (series)
                            VALUES
                            (@series)")
        If SQL.HasException(True) Then Exit Sub
    End Sub
    '' ----------------------------------------------------------------------------------------------------------------
    '' ----------------------------------------------------------------------------------------------------------------
    Public Sub addrelseries(series As String)
        SQL.AddParam("@series", series)

        SQL.ExecQueryDT("INSERT INTO tblrelseries
                            (series)
                            VALUES
                            (@series)")
        If SQL.HasException(True) Then Exit Sub
    End Sub

    Public Sub fetchSeriesRelAndDisplay()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where  receiver != 10006 and  series IS NOT NULL order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer
            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")
                With releaseGoods
                    .series1 = incrementedSeries
                End With

            End If
        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub


    '' ----------------------------------------------------------------------------------------------------------------
    '' ----------------------------------------------------------------------------------------------------------------
    Public Sub addrelPurchaseRetirnseries(series As String)
        SQL.AddParam("@series", series)

        SQL.ExecQueryDT("INSERT INTO tblRelretrunseries
                            (series)
                            VALUES
                            (@series)")
        If SQL.HasException(True) Then Exit Sub
    End Sub


    Public Sub fetchrelPurchaseRetirnseries()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRelsHeader where receiver = 10006 and series is not null  order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer

            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = "1-" & numericPart.ToString("00000")
                With releaseGoods
                    .series1 = incrementedSeries
                End With

                With Print2
                    .series2 = incrementedSeries
                End With
            End If
        Else

            Dim defaultSeries As String = "1-00001"

            With releaseGoods
                .series1 = defaultSeries
            End With

        End If
    End Sub

    Public Sub fetchrECVsUPPLIERseries()
        ' Execute the SQL query to retrieve the top series from tblrecvseries
        SQL.ExecQueryDT("Select TOP 1 series from tblRecvHeaders where sender =10010 and  series is not null order by id desc")

        ' Check for any exceptions during the SQL query execution
        If SQL.HasException(True) Then Exit Sub

        ' Check if there is a record returned from the query
        If SQL.RecordCountDT <> 0 Then
            ' Retrieve the series value from the first row of the result
            Dim originalSeries As String
            originalSeries = SQL.DBDT.Rows(0)("series").ToString()

            ' Extract the numeric part from the series value
            Dim numericPart As Integer
            If Integer.TryParse(originalSeries.Split("-"c)(1), numericPart) Then
                ' Increment the numeric part by 1
                numericPart += 1

                ' Format the incremented series value back into the original format
                Dim incrementedSeries As String = newHome.branch + "-" + numericPart.ToString("00000")

                With recvGoodsMain
                    .series = incrementedSeries
                End With

                'With Print1
                '    .series2 = incrementedSeries
                'End With
            End If
        Else

            Dim defaultSeries As String = newHome.branch + "-00001"

            With recvGoodsMain
                .series = defaultSeries
            End With
        End If
    End Sub

End Class

