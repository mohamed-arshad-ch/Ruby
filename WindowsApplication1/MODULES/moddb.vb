Imports System.Data.OleDb
Module moddb
    Public server As String
    Public Serveruser As String
    Public password As String
    Public dbname As String

    Public Sub setCOnfig()
        Dim FILE_NAME As String = Application.StartupPath & "\startup.ini"

        If System.IO.File.Exists(FILE_NAME) = True Then

            Dim objReader As New System.IO.StreamReader(FILE_NAME)

            server = objReader.ReadLine()
            Serveruser = objReader.ReadLine() & vbNewLine
            password = objReader.ReadLine() & vbNewLine & vbNewLine
            dbname = objReader.ReadLine() & vbNewLine & vbNewLine & vbNewLine


        Else

            MsgBox("Config File Does Not Exist!")

        End If
    End Sub
    Public Function myconn() As OleDbConnection

        setCOnfig()
        Return New OleDbConnection("Provider=" & server & ";Data Source=" & Serveruser & "")

    End Function


End Module

