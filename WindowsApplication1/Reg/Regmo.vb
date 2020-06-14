Imports System.Data.OleDb
Module Regmo
    Public con As OleDbConnection = myconn()
    Public result As Integer
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public dReader As OleDbDataReader
    Public ds As New DataSet
    Public dt As New DataTable
    Public Sub loadaccount(ByVal sql1 As Object, ByVal member As String)


        Try

            Dim id_tmp As Integer
            con.Open()



            cmd = New OleDbCommand(sql1, con)
            dReader = cmd.ExecuteReader
            If dReader.HasRows = False Then
                dReader.Close()
                id_tmp = 200
            Else
                dReader.Read()
                id_tmp = dReader(member) + 1
            End If
            dReader.Close()
            CustomerRegistrationsma.TextBox5.Text = id_tmp
        Catch ex As Exception


        Finally

            con.Close()

        End Try


    End Sub
End Module
