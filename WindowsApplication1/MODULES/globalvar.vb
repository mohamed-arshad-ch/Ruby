Imports System.Data.OleDb
Module globalvar


    Public con As OleDbConnection = myconn()
    Public result As Integer
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public dReader As OleDbDataReader
    Public ds As New DataSet
    Public dt As New DataTable



    Public issucess As Boolean
    Public GLOBALMessage As String = ""


    Public Sub findthis(ByVal sql As String)
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql

            End With
        Catch ex As Exception

        Finally
            con.Close()
            da.Dispose()
        End Try


    End Sub
    



  



    Public Function insert(ByVal sql As String) As Boolean

        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql


                result = cmd.ExecuteNonQuery


                If result = 0 Then
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation)

        Finally
            con.Close()

        End Try
        Return True


    End Function


    Public Function updatefg(ByVal sql As String) As Boolean
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql

                result = cmd.ExecuteNonQuery
                If result = 0 Then

                    Return False
                Else
                    Return True

                End If
            End With
        Catch ex As Exception

            MsgBox(ex.Message)
        Finally
            con.Close()

        End Try
        Return True
    End Function


    Public Function delete(ByVal sql As String) As Boolean
        Try
            con.Open()
            With cmd
                .Connection = con
                .CommandText = sql
                result = cmd.ExecuteNonQuery
                If result = 0 Then
                    Return False
                Else
                    Return True
                End If
            End With
        Catch ex As Exception

        Finally
            con.Close()

        End Try
        Return True
    End Function


    Public Sub loadlist(ByVal sql As Object, ByVal list As Object, ByVal member As String)
        Try

            list.Items.Clear()



            con.Open()

            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader
            list.Items.Clear()
            Dim x As ListViewItem

            Do While dReader.Read = True
                x = New ListViewItem(dReader(member).ToString)
                list.Items.Add(x)

            Loop


        Catch ex As Exception


        Finally


            con.Close()

        End Try

    End Sub




    Public Function NumRows() As Integer
        Try
            con.Open()
            dReader = cmd.ExecuteReader()
            Do While dReader.Read = True

                Return dReader(0)
            Loop

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            con.Close()

        End Try
        Return True

    End Function

  
 



   


    Public Sub slnoforrent(ByVal sql As String)
        Try
            con.Open()





            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader

            If dReader.Read = True Then
                Form1.NumericUpDown1.Text = Val(dReader(0)) + 1
            Else
                Form1.NumericUpDown1.Text = 1
            End If
            dReader.Close()

        Catch ex As Exception


        Finally

            con.Close()

        End Try

    End Sub



    Public Sub slnoforreceipt(ByVal sql As String)
        Try
            con.Open()





            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader

            If dReader.Read = True Then
                Receipt.NumericUpDown1.Text = Val(dReader(0)) + 1
            Else
                Receipt.NumericUpDown1.Text = 1
            End If
            dReader.Close()

        Catch ex As Exception


        Finally

            con.Close()

        End Try

    End Sub


    Public Sub LoadReceiptR(ByVal obj As Object)
        Try
            con.Open()
            dReader = cmd.ExecuteReader()
            '  obj.Rows.Clear()
            Dim x As ListViewItem
            Dim y As Double
            y = 0
            obj.Items.Clear()
            Do While dReader.Read = True
                y = y + 1
                x = New ListViewItem(y)
                x.SubItems.Add(dReader(0).ToString)
                x.SubItems.Add(dReader(1))
                x.SubItems.Add(dReader(2))
                x.SubItems.Add(dReader(3).ToString)
                x.SubItems.Add(dReader(4).ToString)
                x.SubItems.Add(Val(dReader(5)))
                x.SubItems.Add(dReader(6).ToString)

                obj.Items.Add(x)
            Loop


        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            da.Dispose()
            con.Close()
        End Try
    End Sub

End Module
