Public Class Print2
    Public pr As Integer

    Public series2 As String = ""
    Private q As New qry
    Private Sub Print2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReportViewer2.SetDisplayMode(Microsoft.Reporting.WinForms.DisplayMode.PrintLayout)

        ' Set the page settings for portrait orientation with 0.5-inch margins
        Dim pageSettings As New System.Drawing.Printing.PageSettings()
        pageSettings.Landscape = False ' Set to True for landscape orientation
        pageSettings.Margins = New System.Drawing.Printing.Margins(25, 25, 25, 25) ' 0.5 inches on all sides
        ReportViewer2.SetPageSettings(pageSettings)

        ReportViewer2.ZoomMode = Microsoft.Reporting.WinForms.ZoomMode.Percent
        ReportViewer2.ZoomPercent = 100
        ' Optionally, you can refresh the report if needed
        ReportViewer2.RefreshReport()

    End Sub

    Private Sub ReportViewer2_Print(sender As Object, e As Microsoft.Reporting.WinForms.ReportPrintEventArgs) Handles ReportViewer2.Print

    End Sub

    Private Sub Print2_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Me.ReportViewer2.RefreshReport()
        Me.Dispose()
    End Sub
End Class