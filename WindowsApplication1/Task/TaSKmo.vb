Imports System.Data.OleDb
Module TaSKmo
    Public con As OleDbConnection = myconn()
    Public result As Integer
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public dReader As OleDbDataReader
    Public ds As New DataSet
    Public dt As New DataTable
    Public sql As String

   
    Public Sub loadcusto(ByVal sql As String, ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal param4 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True
                Form1.TextBox4.Text = dReader(param1).ToString
                Form1.TextBox5.Text = dReader(param2).ToString
                Form1.TextBox6.Text = dReader(param3).ToString
                Form1.TextBox1.Text = dReader(param4).ToString
            End While
        Catch ex As Exception
        Finally
            con.Close()

        End Try
    End Sub


    Public Sub loadrent(ByVal sql As String, ByVal param1 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True
                Form1.TextBox9.Text = dReader(param1).ToString

            End While
        Catch ex As Exception


        Finally
            con.Close()

        End Try
    End Sub


    Public Sub loadpacka(ByVal sql As String, ByVal param1 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True
                Form1.TextBox11.Text = dReader(param1).ToString

            End While
        Catch ex As Exception


        Finally
            con.Close()

        End Try
    End Sub


    Public Sub combofill(ByVal com As Object, ByVal sql13 As String, ByVal par As String)
        Try


            con.Open()
            com.Items.Clear()
            cmd.CommandText = sql13
            cmd.Connection = con
            dReader = cmd.ExecuteReader
            While dReader.Read
                com.Items.Add(dReader(par))

            End While
            dReader.Close()
            con.Close()
        Catch ex As Exception

        Finally
            con.Close()

        End Try


    End Sub

    Public Sub loadcustoreceipt(ByVal sql As String, ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal param4 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True
                Receipt.TextBox4.Text = dReader(param1).ToString
                Receipt.TextBox5.Text = dReader(param2).ToString
                Receipt.TextBox6.Text = dReader(param3).ToString
                Receipt.TextBox1.Text = dReader(param4).ToString
            End While
        Catch ex As Exception
        Finally
            con.Close()

        End Try
    End Sub


    Public Sub loadcustoreceiptgf(ByVal sql As String, ByVal param1 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True
               
                Receipt.TextBox1.Text = dReader(param1).ToString
            End While
        Catch ex As Exception
        Finally
            con.Close()

        End Try
    End Sub



    Public Sub balanceinre(ByVal sql As String, ByVal param1 As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
            End With
            dReader = cmd.ExecuteReader


            While dReader.Read = True

                Receipt.TextBox10.Text = dReader(param1)
                Receipt.Label6.Text = dReader(param1)



            End While

          
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()

        End Try
    End Sub
    

   
    
    
End Module
