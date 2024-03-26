Imports System.Data.OleDb
Imports System.Data.SqlClient


Public Class addnewitemlist
    Private Sub Addnewitemlist_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using excelConnection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';")
                excelConnection.Open()

                Dim selectQuery As String = "SELECT * FROM [Sheet1$]"

                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using excelReader As OleDbDataReader = excelCommand.ExecuteReader()
                        ' Iterate over the records and create SQL INSERT statements
                        While excelReader.Read()
                            Dim checkQuery As String = "SELECT COUNT(*) FROM tblGoods WHERE id = @Column1"
                            Using checkCommand As New SqlCommand(checkQuery, connection)
                                checkCommand.Parameters.AddWithValue("@Column1", excelReader("id"))
                                Dim count As Integer = Convert.ToInt32(checkCommand.ExecuteScalar())
                                If count = 0 Then
                                    Dim insertQuery As String = "INSERT INTO tblGoods (id, goodDes,LeadTime,LeadTimeAllowance,Status)" &
                                                "VALUES (@Column1, @Column2,@Column3,@Column4,@Column5)"

                                    Using insertCommand As New SqlCommand(insertQuery, connection)
                                        ' Set the parameter values based on your Excel file columns
                                        insertCommand.Parameters.AddWithValue("@Column1", excelReader("id"))
                                        insertCommand.Parameters.AddWithValue("@Column2", excelReader("goodDes"))
                                        insertCommand.Parameters.AddWithValue("@Column3", excelReader("leadtime"))
                                        insertCommand.Parameters.AddWithValue("@Column4", excelReader("leadtimeallowance"))
                                        insertCommand.Parameters.AddWithValue("@Column5", excelReader("status"))

                                        ' Execute the INSERT statement
                                        insertCommand.ExecuteNonQuery()
                                    End Using
                                End If
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

    Private Sub Addnewitemlist_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class