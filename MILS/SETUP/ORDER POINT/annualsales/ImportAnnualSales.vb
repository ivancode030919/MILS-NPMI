Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class ImportAnnualSales
    Private s As New sqlcon
    Dim i As Integer
    '--------------------------------------------------------------------------------------------------------------------------------------
    'import Cash
    '--------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        openFileDialog.Title = "Select an Excel File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbfFilePath As String = openFileDialog.FileName
            TextBox1.Text = dbfFilePath
        End If
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
        InsertExcelDataIntoSQLcash(sqlConnectionString, excelFilePath)
        MessageBox.Show("Successfuly Imported in Database")

    End Sub


    Private Sub InsertExcelDataIntoSQLcash(connectionString As String, excelFilePath As String)
        ' Establish connection to the SQL database
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using excelConnection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';")
                excelConnection.Open()

                Dim selectQuery As String = "SELECT * FROM [Sheet1$]"

                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using excelReader As OleDbDataReader = excelCommand.ExecuteReader()
                        ' Iterate over the records and create SQL INSERT statements
                        While excelReader.Read()
                            Dim insertQuery As String = "INSERT INTO tblCash (goodId, qty) " &
                                "VALUES (@Column1, @Column2)"

                            Using insertCommand As New SqlCommand(insertQuery, connection)
                                ' Set the parameter values based on your Excel file columns
                                insertCommand.Parameters.AddWithValue("@Column1", excelReader("goodId"))
                                insertCommand.Parameters.AddWithValue("@Column2", excelReader("qty"))

                                ' Execute the INSERT statement
                                insertCommand.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                End Using
                excelConnection.Close()
            End Using
            connection.Close()
        End Using
    End Sub
    '------------------------------------------------------------------------------------------------------------------------------------------
    'import Charge Invoice
    '------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        openFileDialog.Title = "Select an Excel File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbfFilePath As String = openFileDialog.FileName
            TextBox2.Text = dbfFilePath
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim excelFilePath As String = TextBox2.Text

        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Please select an Excel file.")
            Return
        End If

        Timer1.Enabled = False

        i = 1

        ' SQL Server connection string
        'Dim sqlConnectionString As String = "Data Source=172.16.10.218;Initial Catalog=New_Dummy_GSC;User ID=HSDPI;Password=123456$hsdp"
        Dim connection As New sqlcon()
        Dim sqlConnectionString As String = connection.DBCon.ConnectionString
        ' Name of the SQL Server table to import data into
        InsertExcelDataIntoSQLCI(sqlConnectionString, excelFilePath)
        MessageBox.Show("Successfuly Imported in Database")




    End Sub


    Private Sub InsertExcelDataIntoSQLCI(connectionString As String, excelFilePath As String)
        ' Establish connection to the SQL database
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using excelConnection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';")
                excelConnection.Open()

                Dim selectQuery As String = "SELECT * FROM [Sheet1$]"

                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using excelReader As OleDbDataReader = excelCommand.ExecuteReader()
                        ' Iterate over the records and create SQL INSERT statements
                        While excelReader.Read()
                            Dim insertQuery As String = "INSERT INTO tblChargeInvoice (goodId, qty) " &
                                "VALUES (@Column1, @Column2)"

                            Using insertCommand As New SqlCommand(insertQuery, connection)
                                ' Set the parameter values based on your Excel file columns
                                insertCommand.Parameters.AddWithValue("@Column1", excelReader("goodId"))
                                insertCommand.Parameters.AddWithValue("@Column2", excelReader("qty"))

                                ' Execute the INSERT statement
                                insertCommand.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                End Using
                excelConnection.Close()
            End Using
            connection.Close()
        End Using
    End Sub

    Private Sub ImportAnnualSales_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    '------------------------------------------------------------------------------------------------------------------------------------------
    'import CM
    '------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If ProgressBar3.Value <= ProgressBar3.Maximum Then
            ProgressBar3.Value += 1
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, ProgressBar3.BackColorChanged
        Dim excelFilePath As String = TextBox3.Text

        If String.IsNullOrEmpty(excelFilePath) Then
            MessageBox.Show("Please select an Excel file.")
            Return
        End If

        Timer1.Enabled = False

        i = 1

        ' SQL Server connection string
        'Dim sqlConnectionString As String = "Data Source=172.16.10.218;Initial Catalog=New_Dummy_GSC;User ID=HSDPI;Password=123456$hsdp"
        Dim connection As New sqlcon()
        Dim sqlConnectionString As String = connection.DBCon.ConnectionString
        ' Name of the SQL Server table to import data into
        InsertExcelDataIntoSQLCM(sqlConnectionString, excelFilePath)
        MessageBox.Show("Successfuly Imported in Database")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim openFileDialog As New OpenFileDialog()
        openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
        openFileDialog.Title = "Select an Excel File"

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim dbfFilePath As String = openFileDialog.FileName
            TextBox3.Text = dbfFilePath
        End If
    End Sub


    Private Sub InsertExcelDataIntoSQLCM(connectionString As String, excelFilePath As String)
        ' Establish connection to the SQL database
        Using connection As New SqlConnection(connectionString)
            connection.Open()

            Using excelConnection As New OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties='Excel 12.0 Xml;HDR=YES';")
                excelConnection.Open()

                Dim selectQuery As String = "SELECT * FROM [Sheet1$]"

                Using excelCommand As New OleDbCommand(selectQuery, excelConnection)
                    Using excelReader As OleDbDataReader = excelCommand.ExecuteReader()
                        ' Iterate over the records and create SQL INSERT statements
                        While excelReader.Read()
                            Dim insertQuery As String = "INSERT INTO tblCm (goodId, qty) " &
                                "VALUES (@Column1, @Column2)"

                            Using insertCommand As New SqlCommand(insertQuery, connection)
                                ' Set the parameter values based on your Excel file columns
                                insertCommand.Parameters.AddWithValue("@Column1", excelReader("goodId"))
                                insertCommand.Parameters.AddWithValue("@Column2", excelReader("qty"))

                                ' Execute the INSERT statement
                                insertCommand.ExecuteNonQuery()
                            End Using
                        End While
                    End Using
                End Using
                excelConnection.Close()
            End Using
            connection.Close()
        End Using
    End Sub

    Private Sub ImportAnnualSales_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class