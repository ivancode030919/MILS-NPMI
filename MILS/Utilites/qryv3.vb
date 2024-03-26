
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

Public Class qryv3
    Public SQL As New sqlcon
    Private f As New functions
    Private recbatch As String

    '=================================================================================================================================================================================
    'UPDATE RECEIVE DETAILS FETCH DOCUMENT TYPE
    '=================================================================================================================================================================================
    Public Sub doctype(ByVal t1 As String)

        SQL.AddParam("@docname", t1)

        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes WHERE remarks =@docname")
        If SQL.HasException(True) Then Exit Sub

        With recvListingDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docTypeId = r(0)

                Next

            End If
        End With


        With releaseDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .doctype = r(0)
                Next

            End If
        End With


    End Sub
    '=================================================================================================================================================================================
    'UPDATE RECEIVE DETAILS FETCH DOCUMENT TYPE REFERENCE
    '=================================================================================================================================================================================
    Public Sub doctype2(ByVal t1 As String)

        SQL.AddParam("@docname", t1)
        SQL.ExecQueryDT("SELECT id FROM tblDocumentTypes WHERE remarks =@docname")
        If SQL.HasException(True) Then Exit Sub
        With recvListingDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .docRefTypeId = r(0)
                Next
            End If
        End With

        With releaseDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .refdoctype = r(0)
                Next
            End If
        End With
    End Sub

    '=================================================================================================================================================================================
    'UPDATE RECEIVE DETAILS FETCH BATCH ID
    '=================================================================================================================================================================================
    Public Sub batchIdsss()

        With releaseDetails.dg1
            For i As Integer = 0 To .Rows.Count - 1
                SQL.AddParam("@batch", .Rows(i).Cells(3).Value)
                SQL.ExecQueryDT("SELECT id FROM tblBatches WHERE batchName=@batch")

                With releaseDetails
                    If SQL.RecordCountDT <> 0 Then
                        For Each r As DataRow In SQL.DBDT.Rows
                            .batch = r(0)

                        Next
                    End If
                End With


            Next
        End With



    End Sub
    '=================================================================================================================================================================================
    'UPDATE RECEIVE DETAILS FETCH LOCATION ID
    '=================================================================================================================================================================================
    Public Sub locidsss()

        With releaseDetails.dg1

            For i As Integer = 0 To .Rows.Count - 1

                SQL.AddParam("@location", .Rows(i).Cells(4).Value)
                SQL.ExecQueryDT("select id from tblLocations where name = @location")


                With releaseDetails
                    If SQL.RecordCountDT <> 0 Then
                        For Each r As DataRow In SQL.DBDT.Rows
                            .loc = r(0)
                        Next
                    End If
                End With
            Next

        End With


    End Sub

    '=================================================================================================================================================================================
    'UPDATE RECEIVE DETAILS ACTION -only SPU
    '=================================================================================================================================================================================
    Public Sub recvupdate(id As String, t1 As String, t2 As String, t3 As String, t4 As Date, t5 As String, t6 As String, t7 As String, t8 As String)

        SQL.AddParam("@id", id)
        SQL.AddParam("@sId", t1)
        SQL.AddParam("@docType", t2)
        SQL.AddParam("@docNum", t3)
        SQL.AddParam("@docDate", t4)
        SQL.AddParam("@rmks", t5)
        SQL.AddParam("@docRefType", t6)
        SQL.AddParam("@docRefNum", t7)
        SQL.AddParam("@oType", t8)



        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double
        Dim xdate As String = ""



        SQL.ExecQueryDT("update tblRecvHeaders set 
                        docType =@docType,docNo= @docNum,docDate= @docDate,sender=@sId,
                        remarks=@rmks,ownershipType=@oType,docRefType=@docRefType,docRefNo= @docRefNum
                        where id = @id")
        If SQL.HasException(True) Then Exit Sub


        SQL.AddParam("@entryId", id)
        SQL.ExecQueryDT("SELECT id,transId,goodId,areaId,locId,batchId,qty,COALESCE(Expirationdate, 'N/A') AS Expirationdate FROM tblRecvDetails
                         wHERE transId = @entryId")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                goodId = r("goodId")
                batchId = r("batchId")
                locId = r("locId")
                qty = r("qty")
                xdate = r("Expirationdate")

                SQL.AddParam("@gId", goodId)
                SQL.AddParam("@bId", batchId)
                SQL.AddParam("@lId", locId)
                SQL.AddParam("@exdate", xdate)

                SQL.ExecQueryDT("SELECT * FROM tblProducts
                                            WHERE goodId = @gId
                                            AND batchId = @bId
                                            AND locationId = @lId                        
                                            AND Expirationdate = @exdate")

                If SQL.HasException(True) Then Exit Sub

                If SQL.RecordCountDT > 0 Then
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)

                    SQL.ExecQueryDT("UPDATE tblProducts
                        set qty = qty - @qty
                        WHERE goodId = @gId
                        AND batchId = @bId
                        AND locationId = @lId
                        AND Expirationdate = @exdate")
                    If SQL.HasException(True) Then Exit Sub
                End If

            Next


        End If



        With recvListingDetails.dg1
            For i As Integer = 0 To .Rows.Count - 1
                SQL.AddParam("@id", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@gId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@loc", .Rows(i).Cells(7).Value.ToString)
                SQL.AddParam("@batch", .Rows(i).Cells(8).Value.ToString)
                SQL.AddParam("@qty", .Rows(i).Cells(5).Value.ToString)
                SQL.AddParam("@exdate", .Rows(i).Cells(6).Value.ToString)

                SQL.ExecQueryDT("UPDATE tblRecvDetails SET goodId=@gId, locId=@loc,batchId=@batch, qty=@qty, Expirationdate=@exdate WHERE id=@id")

                If SQL.HasException(True) Then
                    Exit Sub
                End If
            Next
        End With


        SQL.AddParam("@entryId", id)
        SQL.ExecQueryDT("SELECT * FROM tblRecvDetails
                        WHERE transId = @entryId;")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                goodId = r("goodId")
                batchId = r("batchId")
                locId = r("locId")
                qty = r("qty")
                xdate = r("Expirationdate")
                areaId = r("areaId")
                SQL.AddParam("@gId", goodId)
                SQL.AddParam("@bId", batchId)
                SQL.AddParam("@lId", locId)
                SQL.AddParam("@qty", qty)
                SQL.AddParam("@exdate", xdate)

                SQL.ExecQueryDT("SELECT * FROM tblProducts
                                    WHERE goodId = @gId
                                    AND batchId = @bId
                                    AND locationId = @lId
                                    AND Expirationdate = @exdate")
                If SQL.HasException(True) Then Exit Sub
                If SQL.RecordCountDT > 0 Then
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)

                    SQL.ExecQueryDT("UPDATE tblProducts
                               set qty = qty + @qty
                                WHERE goodId = @gId
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
        MessageBox.Show("Successfully Updated...")

    End Sub

    '=================================================================================================================================================================================
    'INVENTORY GOODS UPDATE FOR STATUS
    '=================================================================================================================================================================================

    Public Sub FetchStatus(t1 As String)
        Dim g As String
        SQL.AddParam("@stat", t1)
        SQL.ExecQueryDT("SELECT Status FROM tblGoods WHERE id =@stat")
        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then

            g = SQL.DBDT.Rows(0)("Status").ToString
            With itemstatus
                If g = 1 Then
                    .RadioButton1.Checked = True
                    .RadioButton2.Checked = False
                    .RadioButton3.Checked = False
                ElseIf g = 2 Then
                    .RadioButton1.Checked = False
                    .RadioButton2.Checked = True
                    .RadioButton3.Checked = False
                ElseIf g = 3 Then
                    .RadioButton1.Checked = False
                    .RadioButton2.Checked = False
                    .RadioButton3.Checked = True
                End If


            End With
        End If
    End Sub
    '=================================================================================================================================================================================
    'UPDATE GOODS STATUS ACTION
    '=================================================================================================================================================================================

    Public Sub updateGoodStatus(t1 As String, t2 As String)

        Dim x As String = t1
        Dim g As String
        Dim h As String

        SQL.AddParam("@id", t2)
        SQL.ExecQueryDT("Select SUM(t1.qty) as qty, t2.Status as stat
                        from tblProducts t1
                        inner join tblGoods t2 on t1.goodId=t2.id
                        where goodId =@id
                        group by t2.Status")
        If SQL.HasException(True) Then Exit Sub
        g = SQL.DBDT.Rows(0)("qty").ToString
        h = SQL.DBDT.Rows(0)("stat").ToString

        If SQL.RecordCountDT <> 0 Then

            If x = "3" Then

                SQL.AddParam("@stat", t1)
                SQL.AddParam("@id", t2)
                SQL.ExecQueryDT("UPDATE tblGoods SET Status = @stat where id= @id")
                If SQL.HasException(True) Then Exit Sub
                MessageBox.Show("Successfully Updated...")


            ElseIf x = "1" Then
                SQL.AddParam("@stat", t1)
                SQL.AddParam("@id", t2)
                SQL.ExecQueryDT("UPDATE tblGoods SET Status = @stat where id= @id")
                If SQL.HasException(True) Then Exit Sub
                MessageBox.Show("Successfully Updated...")
            ElseIf x = "2" Then

                If h = "1" Then
                    If g > 0 Then
                        MessageBox.Show("A product more than 0 may not be updated...")
                    ElseIf g = 0 Then
                        SQL.AddParam("@stat", t1)
                        SQL.AddParam("@id", t2)
                        SQL.ExecQueryDT("UPDATE tblGoods SET Status = @stat where id= @id")
                        If SQL.HasException(True) Then Exit Sub
                        MessageBox.Show("Successfully Updated...")

                    End If

                ElseIf h = "3" Then
                    If g > 0 Then
                        MessageBox.Show("A product more than 0 may not be updated...")
                    Else
                        SQL.AddParam("@stat", t1)
                        SQL.AddParam("@id", t2)
                        SQL.ExecQueryDT("UPDATE tblGoods SET Status = @stat where id= @id")
                        If SQL.HasException(True) Then Exit Sub
                        MessageBox.Show("Successfully Updated...")
                    End If

                Else
                    SQL.AddParam("@stat", t1)
                    SQL.AddParam("@id", t2)
                    SQL.ExecQueryDT("UPDATE tblGoods SET Status = @stat where id= @id")
                    If SQL.HasException(True) Then Exit Sub
                    MessageBox.Show("Successfully Updated...")
                End If

            End If


        End If
        With inventoryView
            .loadLV()
            .tbxPDNO.Text = ""
            .tbxProdDes.Text = ""
        End With


    End Sub

    '=================================================================================================================================================================================
    'CRYSTAL REPORT CONNECTION FOR GOODS STATUS(ACTIVE, IN-ACTIVE, DISCONTINUED)
    '=================================================================================================================================================================================
    Public Sub generateGoodsStatus(ByVal pd As String _
                              , ByVal des As String _
                              , ByVal stat As String _
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

        crParameterDiscreteValue.Value = stat
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("stat")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

    End Sub


    '=================================================================================================================================================================================
    'UPDATE RELEASE DETAILS ACTION -only SPU
    '=================================================================================================================================================================================
    Public Sub relupdate(id As String, t1 As String, t2 As Date, t3 As String, t4 As String, t5 As String, t6 As String, t7 As String, t8 As String)

        SQL.AddParam("@id", id)
        SQL.AddParam("@docNum", t1)
        SQL.AddParam("@docDate", t2)
        SQL.AddParam("@rmks", t3)
        SQL.AddParam("@docRefNum", t4)
        SQL.AddParam("@oType", t5)
        SQL.AddParam("@docType", t6)
        SQL.AddParam("@receiver", t7)
        SQL.AddParam("@docReftype", t8)

        Dim pld As String = ""
        Dim goodId As String = ""
        Dim batchId As String = ""
        Dim locId As String = ""
        Dim areaId As String = ""
        Dim qty As Double
        Dim xdate As String = ""

        '-----------------------------------------
        'update release header
        '-----------------------------------------
        SQL.ExecQueryDT("UPDATE tblRelsHeader SET
                        docType = @docType,docNo=@docNum,docDate=@docDate,receiver=@receiver,remarks=@rmks,ownershipType=@oType,docRefType=@docReftype,docRefNo=@docRefNum
                        where id = @id")
        If SQL.HasException(True) Then Exit Sub



        '-----------------------------------------
        'Select Data From Release Details
        '-----------------------------------------
        SQL.AddParam("@entryId", id)
        SQL.ExecQueryDT("SELECT pId,goodId,areaId,locId,batchId,qty,COALESCE(Expirationdate, 'N/A') AS Expirationdate FROM tblRelsDetails
                         wHERE transId = @entryId")

        If SQL.HasException(True) Then Exit Sub

        If SQL.RecordCountDT <> 0 Then
            For Each r As DataRow In SQL.DBDT.Rows
                pld = r("pId")
                goodId = r("goodId")
                batchId = r("batchId")
                locId = r("locId")
                qty = r("qty")
                xdate = r("Expirationdate")

                SQL.AddParam("@pId", pld)
                SQL.AddParam("@gId", goodId)
                SQL.AddParam("@bId", batchId)
                SQL.AddParam("@lId", locId)
                SQL.AddParam("@exdate", xdate)

                SQL.ExecQueryDT("SELECT * FROM tblProducts
                                            WHERE id = @pId
                                            AND goodId = @gId
                                            AND batchId = @bId
                                            AND locationId = @lId                        
                                            AND Expirationdate = @exdate")
                If SQL.HasException(True) Then Exit Sub


                If SQL.RecordCountDT <> 0 Then

                    SQL.AddParam("@pId", pld)
                    SQL.AddParam("@gId", goodId)
                    SQL.AddParam("@bId", batchId)
                    SQL.AddParam("@lId", locId)
                    SQL.AddParam("@qty", qty)
                    SQL.AddParam("@exdate", xdate)

                    SQL.ExecQueryDT("UPDATE tblProducts
                set qty = qty + @qty
                WHERE id = @pId
                AND goodId = @gId
                AND batchId = @bId
                AND locationId = @lId
                AND Expirationdate = @exdate")

                    If SQL.HasException(True) Then Exit Sub
                End If

            Next

        End If


        '-----------------------------------------
        'insert new data
        '-----------------------------------------
        batchIdsss()
        locidsss()
        With releaseDetails.dg1


            Dim loc As Integer = releaseDetails.loc
            Dim Bat As Integer = releaseDetails.batch
            Dim Area As Integer = releaseDetails.area
            For i As Integer = 0 To .Rows.Count - 1
                SQL.AddParam("@transId", id)
                SQL.AddParam("@id", .Rows(i).Cells(0).Value.ToString)
                SQL.AddParam("@gId", .Rows(i).Cells(1).Value.ToString)
                SQL.AddParam("@loc", .Rows(i).Cells(8).Value.ToString)
                SQL.AddParam("@batch", .Rows(i).Cells(9).Value.ToString)
                SQL.AddParam("@area", Area)
                SQL.AddParam("@qty", .Rows(i).Cells(5).Value.ToString)
                SQL.AddParam("@exdate", .Rows(i).Cells(6).Value.ToString)
                SQL.AddParam("@pId", .Rows(i).Cells(7).Value.ToString)

                SQL.ExecQueryDT("UPDATE tblRelsDetails SET pId=@pId, goodId=@gId,areaId=@area, locId=@loc,batchId=@batch, qty=@qty, Expirationdate=@exdate WHERE id=@id")
                If SQL.HasException(True) Then Exit Sub


            Next
        End With




        SQL.AddParam("@entryId", id)
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
        MessageBox.Show("Successfully Updated...")
    End Sub



    '=================================================================================================================================================================================
    'UPDATE RELEASE FETCH RECEIVER
    '=================================================================================================================================================================================
    Public Sub relreceiver(ByVal t1 As String)

        SQL.AddParam("@area", t1)

        SQL.ExecQueryDT("select id from tblAreas where area = @area")
        If SQL.HasException(True) Then Exit Sub
        With releaseDetails
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .receiver = r(0)
                Next

            End If
        End With


    End Sub


    '=================================================================================================================================================================================
    'SAVE SETUP IN TILES SETUP
    '=================================================================================================================================================================================
    Public Sub SaveTileSetup(t2 As String, t3 As String)
        Dim m As String
        m = t2 * t3

        Dim m1 As String
        Dim m2 As String
        'for cm 
        SQL.AddParam("@m1", t2 + "CM")
        SQL.AddParam("@m2", t3 + "CM")
        SQL.AddParam("@m3", m)

        m1 = t2 / 2.5
        m2 = t3 / 2.5

        SQL.AddParam("@r1", m1 + "IN")
        SQL.AddParam("@r2", m2 + "IN")







        SQL.ExecQueryDT("INSERT INTO tblCalculatorTiles(t1,t2,r1,r2,Measure) 
	                     VALUES (@m1, @m2,@r1,@r2, @m3)")
        If SQL.HasException(True) Then Exit Sub



        With Form3
            .TextBox2.Text = ""
            .TextBox3.Text = ""
            MsgBox("Successfully Saved...")
        End With
        Form3.Dispose()

        With TileSetup
            .loaddg()
        End With
    End Sub


    '=================================================================================================================================================================================
    'DISPLAY DATAGRIDVIEW IN TILES SETUP
    '=================================================================================================================================================================================
    Public Sub ViewTileSetup(v1)
        SQL.AddParam("@des", v1)
        SQL.ExecQueryDT("Select * from 
                            (
                            SELECT  tileID,t1+' X '+t2 as CM,r1+ ' X '+ r2 as INch,Measure,t1,t2 FROM tblCalculatorTiles
                            ) as T0              
                            Order by t0.tileID
                            ")
        If SQL.HasException(True) Then Exit Sub

        With TileSetup.dg1
            .DataSource = SQL.DBDT

            .Columns(0).Visible = False


            .Columns(1).HeaderText = "Tiles (cm)"
            .Columns(2).HeaderText = "Tiles (in)"
            .Columns(3).HeaderText = "Area cm²"



            .Columns(1).Width = 160
            .Columns(2).Width = 160
            .Columns(3).Width = 200

            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True

            .Columns(4).Visible = False
            .Columns(5).Visible = False


            .Columns(1).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)
            .Columns(2).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)
            .Columns(3).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)


            .Columns(1).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)
            .Columns(2).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)
            .Columns(3).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        End With

    End Sub


    '=================================================================================================================================================================================
    'DISPLAY DATA IN DATAGRID IN CALCULATOR DG1
    '=================================================================================================================================================================================
    Public Sub ViewTile()

        SQL.ExecQueryDT("Select * from 
                            (
                            SELECT   tileID,t1+' X '+t2 as CM,r1+ ' X '+ r2 as INch,Measure,  FORMAT(0.00, 'N0') AS Quantity FROM tblCalculatorTiles
                            ) as T0               
                            Order by t0.tileID")
        If SQL.HasException(True) Then Exit Sub

        With Form1.dg1
            .DataSource = SQL.DBDT


            .Columns(0).Visible = False
            .Columns(1).HeaderText = "Tiles (cm)"
            .Columns(2).HeaderText = "Tiles (in)"
            .Columns(4).HeaderText = "Quantity"


            .Columns(0).HeaderCell.Style.Font = New Font(.Font.FontFamily, 15, FontStyle.Bold)
            .Columns(1).HeaderCell.Style.Font = New Font(.Font.FontFamily, 15, FontStyle.Bold)
            .Columns(2).HeaderCell.Style.Font = New Font(.Font.FontFamily, 15, FontStyle.Bold)
            .Columns(4).HeaderCell.Style.Font = New Font(.Font.FontFamily, 15, FontStyle.Bold)


            .Columns(0).DefaultCellStyle.Font = New Font(.Font.FontFamily, 15)
            .Columns(1).DefaultCellStyle.Font = New Font(.Font.FontFamily, 15)
            .Columns(2).DefaultCellStyle.Font = New Font(.Font.FontFamily, 15)
            .Columns(4).DefaultCellStyle.Font = New Font(.Font.FontFamily, 15)



            .Columns(1).Width = 185
            .Columns(2).Width = 185
            .Columns(4).Width = 230
            .Columns(3).Visible = False


            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(4).ReadOnly = True

        End With
    End Sub

    '=================================================================================================================================================================================
    'SAVE SETUP IN TILES SETUP
    '=================================================================================================================================================================================

    Public Sub link(v1 As String)

        SQL.AddParam("@des", v1)
        SQL.ExecQueryDT("SELECT T0.goodId,
                              t1.goodDes,
                              t2.area,
                              SUM(T0.qty) as qty
                            FROM tblProducts t0
                            INNER JOIN tblGoods t1 ON t1.id = t0.goodId
                            INNER JOIN tblAreas t2 ON t2.id = t0.areaId
                            INNER JOIN tblBatches t3 ON t3.id = t0.batchId
                            INNER JOIN tblLocations t4 ON t4.id = t0.locationId
                            Where (t1.goodDes LIKE '%' + @des +'%')
                            AND   (T1.goodDes LIKE '%'+'tile' +'%')
                            GROUP BY T0.goodId, t1.goodDes, t2.area
                            HAVING SUM(t0.qty) > 0
                            ORDER BY t0.goodId ASC;")
        If SQL.HasException(True) Then Exit Sub
        With Form1.dg2
            .DataSource = SQL.DBDT
            .Columns(0).HeaderText = "Product Number"
            .Columns(1).HeaderText = "Description"
            .Columns(2).HeaderText = "Area"
            .Columns(3).HeaderText = "Quantity"

            .Columns(0).Width = 170
            .Columns(1).Width = 550
            .Columns(2).Width = 200
            .Columns(3).Width = 200

            .Columns(0).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)
            .Columns(1).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)
            .Columns(2).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)
            .Columns(3).HeaderCell.Style.Font = New Font(.Font.FontFamily, 11, FontStyle.Bold)

            .Columns(0).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)
            .Columns(1).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)
            .Columns(2).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)
            .Columns(3).DefaultCellStyle.Font = New Font(.Font.FontFamily, 11)


            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True

        End With




    End Sub


    '=================================================================================================================================================================================
    'UPDATE SETUP IN TILES SETUP
    '=================================================================================================================================================================================

    Public Sub UpdateTilesSetup(id As String, v1 As String, v2 As String)

        Dim m As String
        m = v1 * v2

        Dim r1 As String
        Dim r2 As String

        SQL.AddParam("@id", id)
        SQL.AddParam("@r1", v1 + "CM")
        SQL.AddParam("@r2", v2 + "CM")
        SQL.AddParam("@m3", m)

        r1 = v1 / 2.5
        r2 = v2 / 2.5

        SQL.AddParam("@f1", r1 + "IN")
        SQL.AddParam("@f2", r2 + "IN")

        SQL.ExecQueryDT("UPdate tblCalculatorTiles
                            set t1=@r1,t2=@r2,r1=@f1,r2=@f2,Measure=@m3
                            Where tileID=@id")
        If SQL.HasException(True) Then Exit Sub




        With Form2

            MsgBox("Successfully Saved...")
            .TextBox3.Text = ""
            .TextBox2.Text = ""
        End With
        Form2.Dispose()

        With TileSetup
            .loaddg()
        End With

    End Sub

    '=================================================================================================================================================================================
    'AGING VIEW SETUP
    '=================================================================================================================================================================================
    Public Sub Aging(v1, v2)
        SQL.AddParam("@id", v1)
        SQL.AddParam("@des", v2)

        SQL.ExecQueryDT("SELECT TOP 50 * from tblMr t1
                         WHERE (t1.Pdno = @id or @id='')
                         AND (t1.Des LIKE '%' + @des+'%' or @des='')
                         ORDER BY Pdno,Date asc")
        If SQL.HasException(True) Then Exit Sub

        With AgingSetup.dg1

            .DataSource = SQL.DBDT

            .Columns(1).HeaderText = "PRODUCT NUMBER"
            .Columns(2).HeaderText = "DESCRIPTION"
            .Columns(3).HeaderText = "QUANTITY"
            .Columns(4).HeaderText = "MR DATE"

            .Columns(0).Visible = False
            .Columns(1).Width = 150
            .Columns(2).Width = 600
            .Columns(3).Width = 150
            .Columns(4).Width = 150


            .Columns(1).ReadOnly = True
            .Columns(2).ReadOnly = True
            .Columns(3).ReadOnly = True
            .Columns(4).ReadOnly = True


            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .ColumnHeadersDefaultCellStyle.Font = New Font(AgingSetup.dg1.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            Next

        End With

    End Sub
    '=================================================================================================================================================================================
    'REPORT FOR AGING
    '=================================================================================================================================================================================
    Public Sub generateExpiration(ByVal proDuctNo As String _
                              , ByVal des As String _
                              , ByVal date1 As String _
                              , ByVal date2 As String _
                              , ByVal cryRpt As ReportDocument)

        Dim crParameterFieldDefinitions As ParameterFieldDefinitions = cryRpt.DataDefinition.ParameterFields
        Dim crParameterFieldDefinition As ParameterFieldDefinition
        Dim crParameterValues As New ParameterValues
        Dim crParameterDiscreteValue As New ParameterDiscreteValue

        ' Set selectedDocno parameter value
        crParameterDiscreteValue.Value = proDuctNo
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

        crParameterDiscreteValue.Value = date1
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("diff")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

        crParameterDiscreteValue.Value = date2
        crParameterFieldDefinition = crParameterFieldDefinitions.Item("diff1")
        crParameterValues = crParameterFieldDefinition.CurrentValues
        crParameterValues.Clear()
        crParameterValues.Add(crParameterDiscreteValue)
        crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)


    End Sub

    '=================================================================================================================================================================================
    'DELETE TBLMR TABLE
    '=================================================================================================================================================================================

    Public Sub DeletetblMr()
        SQL.ExecQueryDT("DELETE FROM tblMr")
        If SQL.HasException(True) Then Exit Sub
        MsgBox("Data Truncated")
    End Sub
    '=================================================================================================================================================================================
    'DELETE TBLMRTV TABLE
    '=================================================================================================================================================================================

    Public Sub DeletetblMrtv()
        SQL.ExecQueryDT("DELETE FROM tblMrtv")
        If SQL.HasException(True) Then Exit Sub
        MsgBox("Data Truncated")
    End Sub

    'View In Vendor
    Public Sub loadLVRecvDetail(dg As DataGridView, id As String)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("Select * from tblvendor where id=@id or @id=''")
        If SQL.HasException(True) Then Exit Sub
        With MainVendor.dg1
            .AllowUserToResizeColumns = False
            .AllowUserToResizeRows = False
            .AllowUserToAddRows = False
            .DataSource = SQL.DBDT
            .ClearSelection()
            .ColumnHeadersHeight = 45
            .RowTemplate.Height = 35
            '.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
            .Columns(0).HeaderText = "ID"
            .Columns(1).HeaderText = "PRODUCT NUMBER"

            .Columns(0).Width = 100
            .Columns(1).Width = 350

            .Columns(0).ReadOnly = True
            .Columns(1).ReadOnly = True

            '.ColumnHeadersDefaultCellStyle.Font = New Font(dgvProducts.ColumnHeadersDefaultCellStyle.Font, FontStyle.Bold)
            '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
            '.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            For Index As Integer = 0 To .ColumnCount - 1
                .Columns(Index).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                .Columns(Index).HeaderCell.Style.Font = New Font(.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9, FontStyle.Bold)
                .Columns(Index).DefaultCellStyle.Font = New Font(.DefaultCellStyle.Font.FontFamily, 9)
                .Columns(Index).SortMode = DataGridViewColumnSortMode.NotSortable
            Next
            .CurrentCell = Nothing
        End With

    End Sub

    'Add New Vendor
    Public Sub SaveVendor(vendor As String)

        SQL.AddParam("@vendor", vendor)

        SQL.ExecQueryDT("INSERT INTO tblvendor(Vendor) 
	                     VALUES (@vendor)")
        If SQL.HasException(True) Then Exit Sub

        With AddandUpdateVendor
            .TextBox1.Text = String.Empty
            MsgBox("Successfully Saved...")
        End With
        AddandUpdateVendor.Dispose()

        With MainVendor
            .loaddg()
        End With
    End Sub

    Public Sub updateGoodStatus(id As Integer, vendor As String)

        SQL.AddParam("@vendor", vendor)
        SQL.AddParam("@id", id)
        SQL.ExecQueryDT("UPDATE tblvendor SET Vendor = @vendor where id= @id")
        If SQL.HasException(True) Then Exit Sub
        MessageBox.Show("Successfully Updated...")

        AddandUpdateVendor.Dispose()
        With MainVendor
            .loaddg()
        End With

    End Sub

    'For Validation in Transfer
    Public Sub Transferid(ByVal id As String)

        SQL.AddParam("@id", id)

        SQL.ExecQueryDT("Select qty from tblProducts where id = @id")
        If SQL.HasException(True) Then Exit Sub
        With intraTransfer
            If SQL.RecordCountDT <> 0 Then
                For Each r As DataRow In SQL.DBDT.Rows
                    .qty = r(0)
                Next
            End If
        End With


    End Sub

End Class
