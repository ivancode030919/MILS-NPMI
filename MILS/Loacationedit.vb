
Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class Loacationedit
    Private Sub Loacationedit_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        openFileDialog.Title = "Select an Excel File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbfFilePath As String = openFileDialog.FileName
            TextBox1.Text = dbfFilePath
        End If
    End Sub


    Private Sub UpdateExcelDataInSQLCash(connectionString As String, excelFilePath As String)
        ' Establish connection to the SQL database
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using excelConnection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';")
                excelConnection.Open()

                Dim selectQuery As String = "SELECT * FROM [Sheet1$]"

                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using excelReader As OleDbDataReader = excelCommand.ExecuteReader()
                        ' Iterate over the records and create SQL UPDATE statements
                        While excelReader.Read()
                            Dim updateQuery As String = "UPDATE tblCash " &
                                "SET qty = @Column2 " &
                                "WHERE goodId = @Column1"

                            Using updateCommand As New SqlCommand(updateQuery, connection)
                                ' Set the parameter values based on your Excel file columns
                                updateCommand.Parameters.AddWithValue("@Column1", excelReader("goodId"))
                                updateCommand.Parameters.AddWithValue("@Column2", excelReader("qty"))

                                ' Execute the UPDATE statement
                                updateCommand.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                End Using

                excelConnection.Close()
            End Using

            connection.Close()
        End Using
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim excelFilePath As String = TextBox1.Text

        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Please select an Excel file.")
            Return
        End If

        ' SQL Server connection string
        'Dim sqlConnectionString As String = "Data Source=172.16.10.218;Initial Catalog=New_Dummy_GSC;User ID=HSDPI;Password=123456$hsdp"

        Dim connection As New sqlcon()
        Dim sqlConnectionString As String = connection.DBCon.ConnectionString
        ' Name of the SQL Server table to import data into
        UpdateExcelDataInSQLCash(sqlConnectionString, excelFilePath)
        MessageBox.Show("Successfuly Imported in Database")
    End Sub
End Class