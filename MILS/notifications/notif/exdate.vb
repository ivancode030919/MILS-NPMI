Public Class exdate
    Private q As New qry
    Public notif1 As Boolean = False
    Public open As Boolean = False
    Private Sub Exdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        dgcount()
    End Sub
    Public Sub New()
        InitializeComponent()
        ' Get the primary screen's working area (excluding taskbar)
        Dim primaryScreen As Screen = Screen.PrimaryScreen
        Dim workingArea As Rectangle = primaryScreen.WorkingArea
        ' Calculate the y-coordinate to position the form in the upper-right corner with a 10-pixel gap
        Dim y As Integer = workingArea.Y + 30
        ' Set the form's location to the upper-right corner with a 10-pixel gap
        Me.Location = New Point(workingArea.Right - Me.Width, y)
    End Sub

    Private Sub Exdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        dgview.Rows.Clear()
    End Sub

    Private Sub dgcount()
        q.ednotif()
        Dim c As Double = dgview.Rows.Count


        'WithnewHome
        '    .NotificationsToolStripMenuItem.Text = "Notifications (" + c.ToString + ")"
        'End With
    End Sub



    Private Sub Dgview_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles dgview.MouseDoubleClick

        'dgview.Rows(0).Cells
        If dgview.Rows(0).Selected Then
            Expirationdatereport.ShowDialog()
        ElseIf dgview.Rows(1).Selected Then
            OPreport.ShowDialog()
        ElseIf dgview.Rows(2).Selected Then
            Aging.ShowDialog()
        End If


    End Sub

    Private Sub Exdate_MouseLeave(sender As Object, e As EventArgs) Handles MyBase.MouseLeave
        Me.Close()
    End Sub

    Private Sub Exdate_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave
        Me.Close()
    End Sub
End Class