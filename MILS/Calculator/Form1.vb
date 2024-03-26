Public Class Form1

    Private n As New qryv3
    Dim des As String
    Private Sub calculate()
        ' Declare variables to store the values from TextBox1 and TextBox2 as Double.
        Dim v1 As Double
        Dim v2 As Double

        ' Try to parse the text in TextBox1 and TextBox2.
        If Double.TryParse(TextBox1.Text, v1) AndAlso Double.TryParse(TextBox2.Text, v2) Then
            ' Both parsing attempts were successful.
            ' Calculate the product of v1 and v2, format it with commas, and assign it to TextBox3.
            TextBox4.Text = String.Format("{0:N}", v1 * v2)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        calculate()
        If ComboBox2.Text = "Floor Area" Then
            validatingrow()
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        calculate()
        If ComboBox2.Text = "Length & Width" Then
            validatingrow()
        End If
    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        Dim v3 As Double

        If Double.TryParse(TextBox1.Text, v3) Then
            TextBox1.Text = String.Format("{0:N}", v3)

        End If
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        Dim v3 As Double

        If Double.TryParse(TextBox2.Text, v3) Then
            TextBox2.Text = String.Format("{0:N}", v3)

        End If
    End Sub


    Sub cbvalue()
        ComboBox2.Items.AddRange({"Floor Area", "Length & Width"})
        ComboBox1.Items.AddRange({"Centimeter", "Foot", "Inches", "Meter"})
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbvalue()
        n.ViewTile()
        loaddetails()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ComboBox1.Items.Clear()
        ComboBox2.Items.Clear()
    End Sub


    Private Sub loaddetails()
        If ComboBox2.Text = "Floor Area" Then

            Label3.Text = "Floor Area :"
            Label4.Visible = False
            Label5.Visible = False
            Panel4.Visible = False
            Panel5.Visible = False
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox4.Text = ""

        ElseIf ComboBox2.Text = "Length & Width" Then

            Label3.Text = "Length :"
            Label4.Visible = True
            Label5.Visible = True
            Panel4.Visible = True
            Panel5.Visible = True
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox4.Text = ""
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        loaddetails()
        validatingrow()
    End Sub

    Private Sub Panel9_Paint(sender As Object, e As PaintEventArgs) Handles Panel9.Paint
        loaddetails()
    End Sub

    Private Sub validatingrow()
        Try
            Dim measure As Double = 0
            Dim qty As Double = 0
            Dim floorarea As Double
            Dim foot As Double
            Dim meter As Double
            Dim inch As Double

            If ComboBox2.Text = "Floor Area" Then

                If ComboBox1.Text = "Centimeter" Then

                    If Double.TryParse(TextBox1.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value) ' Convert the value in column 2 to a Double
                            qty = Math.Ceiling(floorarea / measure) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If

                ElseIf ComboBox1.Text = "Foot" Then

                    If Double.TryParse(TextBox1.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            foot = measure * 0.00107639
                            qty = Math.Ceiling(floorarea / foot) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If

                ElseIf ComboBox1.Text = "Inches" Then

                    If Double.TryParse(TextBox1.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            inch = measure * 0.155
                            qty = Math.Ceiling(floorarea / inch) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If
                ElseIf ComboBox1.Text = "Meter" Then

                    If Double.TryParse(TextBox1.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            meter = measure / 10000
                            qty = Math.Ceiling(floorarea / meter) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If
                End If

            ElseIf ComboBox2.Text = "Length & Width" Then

                If ComboBox1.Text = "Centimeter" Then


                    If Double.TryParse(TextBox4.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value) ' Convert the value in column 2 to a Double
                            qty = Math.Ceiling(floorarea / measure) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If

                ElseIf ComboBox1.Text = "Foot" Then

                    If Double.TryParse(TextBox4.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            foot = measure * 0.00107639
                            qty = Math.Ceiling(floorarea / foot) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If
                ElseIf ComboBox1.Text = "Inches" Then
                    If Double.TryParse(TextBox4.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            inch = measure * 0.155
                            qty = Math.Ceiling(floorarea / inch) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If
                ElseIf ComboBox1.Text = "Meter" Then

                    If Double.TryParse(TextBox4.Text, floorarea) Then
                        For Each row As DataGridViewRow In dg1.Rows
                            measure = CDbl(row.Cells(3).Value)
                            ' Convert the value in column 2 to a Double
                            meter = measure / 10000
                            qty = Math.Ceiling(floorarea / meter) ' Round up the result
                            row.Cells(4).Value = String.Format("{0:F2}", qty)
                        Next
                    End If
                End If

            End If


        Catch ex As Exception
            MsgBox("here")
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        n.ViewTile()
        validatingrow()

    End Sub

    Private Sub Dg1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dg1.CellClick
        If dg1.SelectedRows.Count > 0 Then
            ' Get the selected row
            Dim selectedRow As DataGridViewRow = dg1.SelectedRows(0)

            ' Check if the cell in the first column of the selected row is not null
            If selectedRow.Cells(1).Value IsNot Nothing Then
                ' Retrieve and display the value in a message box
                des = selectedRow.Cells(1).Value.ToString()
                dg2display()
            End If
        End If
    End Sub


    Private Sub dg2display()
        n.link(des)
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
End Class