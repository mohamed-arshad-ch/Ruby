Imports System.Data.OleDb
Public Class RentReportFilter
    Dim con As OleDbConnection = myconn()
    Dim cmd As New OleDbCommand
    Dim da As New OleDbDataAdapter
    Dim dReader As OleDbDataReader
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim sql As String
    Dim agt As New ListView
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub RentReportFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        combofill(ComboBox1, "select code from ChartOfCode where Uid=201", "code")
        combofill(ComboBox2, "select code from ChartOfCode where Uid=202", "code")
    End Sub


    Public Sub ret()
        RentReport.ListView1.Items.Clear()


        sql = "SELECT RentDetails.InvNo, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType ,RentDetails.Discount,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 FROM (RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID )  WHERE RentDetails.RentType='" & ComboBox1.Text & "' " & "AND" & "  RentDetails.PackageType='" & ComboBox2.Text & "'  ORDER BY RentDetails.InvNo "
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
                x.SubItems.Add(dReader("InvNo").ToString)
                x.SubItems.Add(dReader("IDate"))
                x.SubItems.Add(dReader("AcName").ToString)
                x.SubItems.Add(dReader("Address1").ToString)
                x.SubItems.Add(dReader("RentType").ToString)
                x.SubItems.Add(dReader("PackageType").ToString)
                x.SubItems.Add(Val(dReader("Discount")))
                x.SubItems.Add(dReader("NetAmount").ToString)






                RentReport.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()


    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        Select Case Button1.Text = "OK"
            Case ComboBox1.Text.Length > 0 And ComboBox2.Text.Length > 0
                ret()
                RentReport.Show()

            Case ComboBox1.Text.Length > 0

                ret2()
                RentReport.Show()

            Case ComboBox2.Text.Length > 0
                ret3()
                RentReport.Show()


            Case ComboBox1.Text = "" And ComboBox2.Text = ""
                ret1()
                RentReport.Show()




        End Select

    End Sub

    Public Sub ret1()
        RentReport.ListView1.Items.Clear()


        sql = "SELECT RentDetails.InvNo, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType ,RentDetails.Discount,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 FROM (RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID )    ORDER BY RentDetails.InvNo "
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
                x.SubItems.Add(dReader("InvNo").ToString)
                x.SubItems.Add(dReader("IDate"))
                x.SubItems.Add(dReader("AcName").ToString)
                x.SubItems.Add(dReader("Address1").ToString)
                x.SubItems.Add(dReader("RentType").ToString)
                x.SubItems.Add(dReader("PackageType").ToString)
                x.SubItems.Add(Val(dReader("Discount")))
                x.SubItems.Add(dReader("NetAmount").ToString)






                RentReport.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub


    Public Sub ret2()
        RentReport.ListView1.Items.Clear()


        sql = "SELECT RentDetails.InvNo, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType ,RentDetails.Discount,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 FROM (RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID ) WHERE RentDetails.RentType='" & ComboBox1.Text & "'   ORDER BY RentDetails.InvNo "
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
                x.SubItems.Add(dReader("InvNo").ToString)
                x.SubItems.Add(dReader("IDate"))
                x.SubItems.Add(dReader("AcName").ToString)
                x.SubItems.Add(dReader("Address1").ToString)
                x.SubItems.Add(dReader("RentType").ToString)
                x.SubItems.Add(dReader("PackageType").ToString)
                x.SubItems.Add(Val(dReader("Discount")))
                x.SubItems.Add(dReader("NetAmount").ToString)






                RentReport.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub


    Public Sub ret3()
        RentReport.ListView1.Items.Clear()


        sql = "SELECT RentDetails.InvNo, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType ,RentDetails.Discount,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 FROM (RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID ) WHERE RentDetails.PackageType='" & ComboBox2.Text & "'   ORDER BY RentDetails.InvNo "
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
                x.SubItems.Add(dReader("InvNo").ToString)
                x.SubItems.Add(dReader("IDate"))
                x.SubItems.Add(dReader("AcName").ToString)
                x.SubItems.Add(dReader("Address1").ToString)
                x.SubItems.Add(dReader("RentType").ToString)
                x.SubItems.Add(dReader("PackageType").ToString)
                x.SubItems.Add(Val(dReader("Discount")))
                x.SubItems.Add(dReader("NetAmount").ToString)






                RentReport.ListView1.Items.Add(x)


            Loop



        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        con.Close()

    End Sub
End Class