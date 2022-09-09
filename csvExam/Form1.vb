Imports Microsoft.VisualBasic.FileIO

Imports System.Data.OleDb
Imports System.IO
Imports System.Text

Public Class Form1
    Private conn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\csvExam\master.mdb")
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim dr2 As OleDbDataReader
    Dim dr3 As OleDbDataReader
    Dim dr4 As OleDbDataReader
    Dim dr5 As OleDbDataReader
    Dim dr6 As OleDbDataReader
    Dim dr7 As OleDbDataReader
    Dim StatusClient As String
    Dim command As String
    Dim dt As New DataTable

    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        DataGridView1.Rows.Clear()
        Try
            Dim fname As String = "c:\csvExam\order_log00.csv"
            Dim reader As New StreamReader(fname, Encoding.Default)
            Dim sline As String = ""
            Dim r As Integer
            sline = reader.ReadLine
            Do
                sline = reader.ReadLine
                If sline Is Nothing Then Exit Do
                Dim words() As String = sline.Split(",")
                DataGridView1.Rows.Add()
                For i As Integer = 0 To 4
                    DataGridView1.Rows(r).Cells(i).Value = words(i)
                Next
                r = r + 1
            Loop
            reader.Close()

        Catch ex As Exception

        End Try
        For i As Integer = 0 To DataGridView1.Rows.Count - 2
            conn.Open()
            cmd = New OleDbCommand("Insert into csvRead (id, area, product, qty, brand, order_date) VALUES(@id, @area, @product, @qty, @brand, @order_date)", conn)
            cmd.Parameters.AddWithValue("@id", DataGridView1.Rows(i).Cells(0).Value.ToString)
            cmd.Parameters.AddWithValue("@area", DataGridView1.Rows(i).Cells(1).Value.ToString)
            cmd.Parameters.AddWithValue("@product", DataGridView1.Rows(i).Cells(2).Value.ToString)
            cmd.Parameters.AddWithValue("@qty", DataGridView1.Rows(i).Cells(3).Value.ToString)
            cmd.Parameters.AddWithValue("@brand", DataGridView1.Rows(i).Cells(4).Value.ToString)
            cmd.Parameters.AddWithValue("@order_date", Format(CDate(Today), "dd/MM/yyyy"))

            cmd.ExecuteNonQuery()
            conn.Close()

        Next
        MsgBox("Data Successfully inserted....")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        conn.Open()
        Dim csvFile As String = "c:\csvExam\0_order_log00.csv"
        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)
        Dim curdate As String
        Dim Cmd2 As New OleDbCommand
        Dim Cmd1 As New OleDbCommand
        Dim Cmd3 As New OleDbCommand
        Dim Cmd4 As New OleDbCommand
        Dim Cmd5 As New OleDbCommand
        Dim total_order As Integer
        Dim product_name As String
        Dim cur_product_name As String = ""
        Dim product_count As Integer = 0
        Cmd2.CommandText = "select * from csvRead"
        Cmd2.Connection = conn
        dr = Cmd2.ExecuteReader
        While dr.Read
            total_order = total_order + 1

        End While

        conn.Close()

        conn.Open()
        curdate = Format(CDate(Today), "dd/MM/yyyy")
        Cmd3.CommandText = "select * from store_product"
        Cmd3.Connection = conn
        dr2 = Cmd3.ExecuteReader

        While dr2.Read
            product_count = 0
            Cmd4.CommandText = "select * from csvRead"
            Cmd4.Connection = conn
            dr3 = Cmd4.ExecuteReader
            While dr3.Read
                If (dr2("Product_name")) = (dr3("product")) Then
                    product_count = product_count + (dr3("qty"))
                    product_name = (dr3("product"))
                End If

            End While

            outFile.Write((dr2("Product_name")))
            outFile.Write((","))
            outFile.WriteLine((product_count / total_order))
            dr3.Close()
        End While
        conn.Close()
        outFile.Close()
        Console.WriteLine(My.Computer.FileSystem.ReadAllText(csvFile))
        MsgBox("0_log csv file Successfully created....")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Cmd6 As New OleDbCommand
        conn.Open()
        Dim csvFile As String = "c:\csvExam\1_order_log00.csv"
        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)
        Cmd6.CommandText = "SELECT TOP 2 brand, product, Sum(qty) as qty FROM csvRead group by brand, product order by SUM(qty) desc"
        Cmd6.Connection = conn
        dr4 = Cmd6.ExecuteReader
        While dr4.Read
            outFile.Write((dr4("Product")))
            outFile.Write((","))
            outFile.WriteLine((dr4("brand")))

        End While
        conn.Close()
        outFile.Close()
        Console.WriteLine(My.Computer.FileSystem.ReadAllText(csvFile))
        MsgBox("1_log csv file Successfully created....")
    End Sub
End Class
