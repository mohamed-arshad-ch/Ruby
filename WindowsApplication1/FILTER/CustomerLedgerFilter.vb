Imports System.Data.OleDb
Public Class CustomerLedgerFilter
    Dim con As OleDbConnection = myconn()
    Dim cmd As New OleDbCommand
    Dim da As New OleDbDataAdapter
    Dim dReader As OleDbDataReader
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim sql As String
    Dim agt As New ListView
    Private Sub LedgerFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        combofill(ComboBox1, "select AcName from ChartOfAccounts where GroupID=100", "AcName")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ComboBox1.Text = "" Then
            ret1()
            CustomerLedger.Show()
        Else
            ret()
            CustomerLedger.Show()
        End If


    End Sub

    Public Sub ret1()
        CustomerLedger.ListView1.Items.Clear()


        sql = "SELECT  ChartOfAccounts.AcName,ChartOfAccounts.Address1,SUM(Debit)-SUM(Credit) AS [Total Quantity] FROM(GeneralJournal INNER JOIN ChartOfAccounts ON GeneralJournal.AccountID=ChartOfAccounts.AccountID) WHERE ChartOfAccounts.GroupID=100  GROUP BY ChartOfAccounts.AcName,ChartOfAccounts.Address1,ChartOfAccounts.GroupID "

        cmd = New OleDbCommand(sql, con)
        Try
            con.Open()
            dReader = cmd.ExecuteReader
            Dim x As ListViewItem
            Dim y As Double
            y = 0
            Do While dReader.Read = True







                y = y + 1

                x = New ListViewItem(y)


                x.SubItems.Add(dReader("AcName"))
                x.SubItems.Add(dReader("Address1"))

                x.SubItems.Add(Val(dReader("Total Quantity")))








                CustomerLedger.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub


    Public Sub ret()
        CustomerLedger.ListView1.Items.Clear()


        sql = "SELECT  ChartOfAccounts.AcName,ChartOfAccounts.Address1,SUM(Debit)-SUM(Credit) AS [Total Quantity] FROM(GeneralJournal INNER JOIN ChartOfAccounts ON GeneralJournal.AccountID=ChartOfAccounts.AccountID) WHERE ChartOfAccounts.AcName='" & ComboBox1.Text & "' GROUP BY ChartOfAccounts.AcName,ChartOfAccounts.Address1 "

        cmd = New OleDbCommand(sql, con)
        Try
            con.Open()
            dReader = cmd.ExecuteReader
            Dim x As ListViewItem
            Dim y As Double
            y = 0
            Do While dReader.Read = True







                y = y + 1

                x = New ListViewItem(y)


                x.SubItems.Add(dReader("AcName"))
                x.SubItems.Add(dReader("Address1"))

                x.SubItems.Add(Val(dReader("Total Quantity")))








                CustomerLedger.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class