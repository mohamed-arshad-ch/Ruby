Imports System.Data.OleDb
Public Class Receipt
    Public con As OleDbConnection = myconn()
    Public result As Integer
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public dReader As OleDbDataReader
    Public ds As New DataSet
    Public dt As New DataTable
    Public sql As String
    Private Sub Receipt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        combofill(ComboBox3, "select AcName from ChartOfAccounts where GroupID=100 ", "AcName")
        slno()
        TextBox10.Text = 0

    End Sub
    Private Sub slno()
        slnoforreceipt("select  RPNo  from ReceiptPayment where TypeCode=15 order by RPNo desc ")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub loadaddressre()
        TextBox10.Text = 0
        Label6.Text = 0
        balanceinre("SELECT  ChartOfAccounts.AcName,ChartOfAccounts.Address1,SUM(Debit)-SUM(Credit) AS [Total Quantity] FROM(GeneralJournal INNER JOIN ChartOfAccounts ON GeneralJournal.AccountID=ChartOfAccounts.AccountID) WHERE ChartOfAccounts.AcName='" & ComboBox3.Text & "' GROUP BY ChartOfAccounts.AcName,ChartOfAccounts.Address1 ", "Total Quantity")
        loadcustoreceiptgf("select AccountID from ChartOfAccounts where Description='" & ComboBox3.Text & "'  ", "AccountID")
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""


        loadcustoreceipt("SELECT PartyDetails.Address2,PartyDetails.Address3,ChartOfAccounts.AccountID,ChartOfAccounts.Address1 FROM(PartyDetails) INNER JOIN ChartOfAccounts ON PartyDetails.AccountID=ChartOfAccounts.AccountID WHERE ChartOfAccounts.Description='" & ComboBox3.Text & "'", "Address1", "Address2", "Address3", "AccountID")

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        loadaddressre()

    End Sub
    Private Sub calcu()
        Dim var1 As Double
        Dim var2 As Double

        Dim net As Double

        var1 = Val(Label6.Text)
        var2 = Val(TextBox8.Text)

        net = (var1 - var2)


        TextBox10.Text = net



    End Sub



    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        calcu()
    End Sub

    Private Sub insertreceipt()
        insert(" insert into ReceiptPayment values ('15','" & NumericUpDown1.Text & "','" & DateTimePicker1.Text & "','" & TextBox1.Text & "','" & TextBox8.Text & "',0,'" & TextBox3.Text & "','" & TextBox10.Text & "')")
        insert("insert into GeneralJournal values ('15','" & DateTimePicker1.Text & "','" & NumericUpDown1.Text & "','" & TextBox1.Text & "','0','" & TextBox8.Text & "','" & TextBox3.Text & "',1)")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        vali()
    End Sub

    Private Sub CLEAR()
        ComboBox3.Text = ""
        TextBox4.Text = ""
        TextBox10.Text = ""
        TextBox1.Text = ""
        TextBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox8.Text = ""
    End Sub

    Private Sub updatereceipt()



        sql = "UPDATE [ReceiptPayment] SET [TypeCode]=15, [TDate]='" & DateTimePicker1.Value & "',[AccountID]='" & TextBox1.Text & "', " & _
                                  "[Amount] ='" & TextBox8.Text & "',[Dis]=0, [Narration] = '" & TextBox3.Text & "' ,[OBalance] ='" & TextBox10.Text & "'  WHERE [RPNo] = " & NumericUpDown1.Value
        issucess = updatefg(sql)

        sql = "UPDATE [GeneralJournal] SET [EDate]='" & DateTimePicker1.Value & "',[AccountID]='" & TextBox1.Text & "',[Debit] =0,[Credit] ='" & TextBox8.Text & "', " & _
                                 "[Narration] ='" & TextBox3.Text & "', [ToAccountID] = 1  WHERE [VoucherNo] = " & NumericUpDown1.Value & " AND GJType=15"
        issucess = updatefg(sql)
        GLOBALMessage = "UpdateOnly"


    End Sub

    Private Sub vali()
        If ComboBox3.Text = "" Or TextBox8.Text = "" Then
            MessageBox.Show("Please Enter Valid Data!", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Try

                findthis("SELECT count(*) FROM ReceiptPayment WHERE RPNo=" & NumericUpDown1.Text & "")
                If NumRows() > 0 Then

                    updatereceipt()
                    CLEAR()
                    slno()
                    TextBox10.Text = ""
                    TextBox8.Text = ""
                Else
                    insertreceipt()
                    CLEAR()
                    slno()
                    TextBox10.Text = ""
                    TextBox8.Text = ""
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
    End Sub


    Private Sub loadrpno()



        Try
            con.Open()

            sql = "SELECT ReceiptPayment.AccountID, ReceiptPayment.TDate,ReceiptPayment.Narration,ReceiptPayment.Dis,ReceiptPayment.OBalance,ReceiptPayment.Amount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 ,PartyDetails.Address2,PartyDetails.Address3 FROM ((ReceiptPayment  INNER JOIN ChartOfAccounts ON ReceiptPayment.AccountID = ChartOfAccounts.AccountID ) INNER JOIN PartyDetails  ON ReceiptPayment.AccountID=PartyDetails.AccountID) WHERE ReceiptPayment.RPNo=" & NumericUpDown1.Text & " "


            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader
            While dReader.Read = True



                TextBox4.Text = dReader("Address1").ToString
                TextBox5.Text = dReader("Address2").ToString
                TextBox6.Text = dReader("Address3").ToString
                TextBox10.Text = dReader("OBalance")
                TextBox3.Text = dReader("Narration").ToString
                TextBox1.Text = dReader("AccountID").ToString

                TextBox8.Text = dReader("Amount").ToString

                ComboBox3.Text = dReader("AcName").ToString


            End While

        Catch ex As OleDbException
            MsgBox(ex.Message)

        Finally
            con.Close()

        End Try
    End Sub

    Private Sub NumericUpDown1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.Click
        JDFGD()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        JDFGD()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        CLEAR()
        slno()
        TextBox10.Text = ""
        TextBox8.Text = ""

    End Sub
    Private Sub JDFGD()
        Try
            findthis("SELECT count(*) FROM ReceiptPayment WHERE RPNo=" & NumericUpDown1.Text & "")
            If NumRows() = 1 Then


                loadrpno()
            Else

                CLEAR()
                TextBox10.Text = ""
                TextBox8.Text = ""

            End If
        Catch ex As Exception

        End Try

    End Sub
    Private Sub deleterent()
        delete("delete from ReceiptPayment where RPNo = " & NumericUpDown1.Value & " ")
        delete("delete from GeneralJournal where VoucherNo = " & NumericUpDown1.Value & " AND GJType=15 ")


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If MessageBox.Show("Sure Delete!", "ALERT", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) = MsgBoxResult.Ok Then
            deleterent()
            CLEAR()
            slno()
            TextBox10.Text = ""
            TextBox8.Text = ""

        End If
    End Sub
End Class