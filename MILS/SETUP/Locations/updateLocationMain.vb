Public Class updateLocationMain
    Public id As String = ""
    Private q As New qry
    Public apprvBy As String = ""

    Private Sub updateLocationMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadDatas()
    End Sub

    Sub loadDatas()
        q.fethcUpdateLocMainDatas(id, newHome.areaId)
    End Sub

    Private Sub updateLocationMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Home_Page.Enabled = True



        If newHome.userId = 20020 Then
            With viewLocations
                .Enabled = True
                .btnDetails.Enabled = False
                .lvLocations.SelectedItems.Clear()
                .Focus()
                .Select()
                .tbxFilter.Text = ""
                .loadLVadmin()
            End With
        Else
            With viewLocations
                .Enabled = True
                .btnDetails.Enabled = False
                .lvLocations.SelectedItems.Clear()
                .Focus()
                .Select()
                .tbxFilter.Text = ""
                .loadLV()
            End With
        End If
    End Sub

    Private Sub updateLocationMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Home_Page.Enabled = True
            With viewLocations
                .Enabled = True
                .btnDetails.Enabled = False
                .lvLocations.SelectedItems.Clear()
                .Focus()
                .Select()
                .tbxFilter.Text = ""
                .loadLV()
            End With
            Me.Dispose()
        Else

        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If String.IsNullOrWhiteSpace(tbxLocName.Text) Or String.IsNullOrWhiteSpace(tbxDescrip.Text) Then
            lblErr.Visible = True
            lblErr.Text = "Fields with * are required."
        Else
            lblErr.Visible = False
            Me.Enabled = False
            updateLocationConfrimation.Show()
        End If
    End Sub

    Sub updateDatas()
        q.updateLocMainQry(id, tbxLocName.Text, tbxDescrip.Text, newHome.userId, newHome.areaId, Date.Now, apprvBy)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class