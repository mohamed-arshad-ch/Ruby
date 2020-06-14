Imports System.Data.OleDb
Public Class Form1
    Public con As OleDbConnection = myconn()
    Public result As Integer
    Public cmd As New OleDbCommand
    Public da As New OleDbDataAdapter
    Public dReader As OleDbDataReader
    Public ds As New DataSet
    Public dt As New DataTable
    Public sql As String


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fillcombobox()
        slno()
        CLear()
        TextBox7.Text = 0
        DateTimePicker1.Value = Today



    End Sub

    Private Sub slno()
        slnoforrent("select  InvNo  from RentDetails order by InvNo desc ")
    End Sub

    Private Sub fillcombobox()
        combofill(ComboBox3, "select AcName from ChartOfAccounts where GroupID=100", "AcName")
        combofill(ComboBox1, "select code from ChartOfCode where Uid=201", "code")
        combofill(ComboBox2, "select code from ChartOfCode where Uid=202", "code")
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged

        loadaddress()


    End Sub

    Private Sub loadaddress()
        loadcusto("SELECT PartyDetails.Address2,PartyDetails.Address3,PartyDetails.AccountID,ChartOfAccounts.Address1 FROM(PartyDetails) INNER JOIN ChartOfAccounts ON PartyDetails.AccountID=ChartOfAccounts.AccountID WHERE ChartOfAccounts.Description='" & ComboBox3.Text & "'", "Address1", "Address2", "Address3", "AccountID")
    End Sub

    Private Sub CLear()
        ComboBox3.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox4.Text = ""
        TextBox1.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        TextBox8.Text = ""
        TextBox10.Text = ""
        TextBox7.Text = ""

        TextBox3.Text = ""


    End Sub

    Private Sub DesighnPrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DesighnPrintToolStripMenuItem.Click
        CLear()
        slno()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        CLear()
        slno()

    End Sub

    Private Sub loadre()
        loadrent("SELECT Description from ChartOfCode WHERE Code='" & ComboBox1.Text & "'", "Description")

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        loadre()
        calcu()
        TextBox7.Text = 0
    End Sub

    Private Sub loadPACKAG()
        loadpacka("SELECT Description from ChartOfCode WHERE Code='" & ComboBox2.Text & "'", "Description")
     
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        loadPACKAG()
        calcu()
        TextBox7.Text = 0
    End Sub

    Private Sub calcu()
        Dim var1 As Double
        Dim var2 As Double
        Dim var3 As Double
        Dim var4 As Double
        Dim net As Double

        var1 = Val(TextBox9.Text)
        var2 = Val(TextBox11.Text)

        var3 = (var1 + var2)
        var4 = Val(TextBox7.Text)

        Net = (var3 - var4)

        TextBox8.Text = Net
        TextBox10.Text = net

    End Sub

    Private Sub TextBox7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox7.TextChanged
        calcu()
    End Sub

    Private Sub insertrent()

        issucess = insert("insert into RentDetails values ('" & NumericUpDown1.Text & "','" & DateTimePicker1.Value & "','" & TextBox1.Text & "','" & ComboBox1.Text & "','" & TextBox9.Text & "','" & ComboBox2.Text & "','" & TextBox11.Text & "','" & TextBox3.Text & "','" & TextBox7.Text & "','" & TextBox10.Text & "','" & TextBox8.Text & "')")
        issucess = insert("insert into GeneralJournal values (30,'" & DateTimePicker1.Text & "','" & NumericUpDown1.Text & "','" & TextBox1.Text & "','" & TextBox8.Text & "',0,'" & TextBox3.Text & "',15)")
        ' issucess = insert("insert into GeneralJournal values (30,'" & DateTimePicker1.Text & "','" & NumericUpDown1.Text & "',1,0,'" & TextBox8.Text & "','" & TextBox3.Text & "','" & TextBox1.Text & "')")

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        vali()
    End Sub


    Private Sub findthiscus()

        Try
            findthis("SELECT count(*) FROM ChartOfAccounts WHERE Description='" & ComboBox3.Text & "'")
            If NumRows() > 0 Then
                loadaddress()
            Else
                CustomerRegistrationsma.Show()
                CustomerRegistrationsma.TextBox1.Text = ComboBox3.Text

                CustomerRegistrationsma.TextBox2.Select()

            End If
        Catch ex As Exception
            MsgBox("asdsdf")
        End Try
      
    End Sub

    Private Sub fd(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            If ComboBox3.Text = "" Then
                MessageBox.Show("Please Enter Valid Data!", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Else
                findthiscus()


            End If

        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        vali()
    End Sub
    Private Sub vali()
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Or ComboBox3.Text = "" Then
            MessageBox.Show("Please Enter Valid Data!", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Try

                findthis("SELECT count(*) FROM RentDetails WHERE InvNo=" & NumericUpDown1.Text & "")
                If NumRows() > 0 Then
                    updaterent()
                    CLear()
                    slno()

                Else
                    insertrent()
                    fillcombobox()
                    CLear()
                    slno()
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub loadsl()

    End Sub

    Private Sub NumericUpDown1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NumericUpDown1.Click
        JDFGD()
    End Sub

    Private Sub NumericUpDown1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NumericUpDown1.ValueChanged
        
        JDFGD()

    End Sub
    Private Sub JDFGD()
        Try
            findthis("SELECT count(*) FROM RentDetails WHERE InvNo=" & NumericUpDown1.Text & "")
            If NumRows() = 1 Then


                loadinvoice()
            Else

                CLear()

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReportToolStripMenuItem.Click


        



    End Sub


    Public Sub loadinvoice()
        Try
            con.Open()

            sql = "SELECT RentDetails.AccountID, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType,RentDetails.Rate0,RentDetails.Rate1,RentDetails.Desc,RentDetails.Discount,RentDetails.Balance,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 ,PartyDetails.Address2,PartyDetails.Address3 FROM ((RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID ) INNER JOIN PartyDetails  ON RentDetails.AccountID=PartyDetails.AccountID) WHERE RentDetails.InvNo=" & NumericUpDown1.Text & " "


            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader
            While dReader.Read = True


                ComboBox2.Text = dReader("PackageType").ToString
                TextBox4.Text = dReader("Address1").ToString
                TextBox5.Text = dReader("Address2").ToString
                TextBox6.Text = dReader("Address3").ToString
                TextBox10.Text = dReader("Balance").ToString
                TextBox3.Text = dReader("Desc").ToString
                TextBox1.Text = dReader("AccountID").ToString
                TextBox9.Text = dReader("Rate0").ToString
                TextBox11.Text = dReader("Rate1").ToString
                ComboBox1.Text = dReader("RentType").ToString
                TextBox8.Text = dReader("Netamount").ToString
                TextBox7.Text = Val(dReader("Discount"))
                ComboBox3.Text = dReader("AcName").ToString


            End While

        Catch ex As OleDbException

        Finally
            con.Close()

        End Try
    End Sub

    Private Sub updaterent()


        sql = "UPDATE [RentDetails] SET [IDate]='" & DateTimePicker1.Value & "',[AccountID]='" & TextBox1.Text & "',[RentType] = '" & ComboBox1.Text & "', [Rate0] ='" & TextBox9.Text & "',[PackageType] ='" & ComboBox2.Text & "', " & _
                                  "[Rate1] ='" & TextBox11.Text & "', [Desc] = '" & TextBox3.Text & "',[Discount] ='" & TextBox7.Text & "' ,[Balance] ='" & TextBox10.Text & "',[NetAmount] ='" & TextBox8.Text & "' WHERE [InvNo] = " & NumericUpDown1.Value
        issucess = updatefg(sql)

        sql = "UPDATE [GeneralJournal] SET [EDate]='" & DateTimePicker1.Value & "',[AccountID]='" & TextBox1.Text & "',[Debit] ='" & TextBox8.Text & "',[Credit] =0, " & _
                                 "[Narration] ='" & TextBox3.Text & "', [ToAccountID] = 15 WHERE [VoucherNo] = " & NumericUpDown1.Value & " AND GJType=30"
        issucess = updatefg(sql)
        GLOBALMessage = "UpdateOnly"
    End Sub

    Private Sub deleterent()
        delete("delete from RentDetails where InvNo = " & NumericUpDown1.Value & " ")
        delete("delete from GeneralJournal where VoucherNo = " & NumericUpDown1.Value & " AND GJType=30 ")


    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If MessageBox.Show("Sure Delete!", "ALERT", MessageBoxButtons.OKCancel, MessageBoxIcon.Error) = MsgBoxResult.Ok Then
            deleterent()
            CLear()
            slno()
        End If


    End Sub

   
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If ComboBox3.Text = "" Then
            MessageBox.Show("Please Enter Valid Data!", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Receiptprint1.Label3.Text = NumericUpDown1.Text
            vali()
            LOADPRIN()
            Receiptprint1.Show()
            


        End If





    End Sub



    Public Sub LOADPRIN()
        Try
            con.Open()

            sql = "SELECT RentDetails.AccountID, RentDetails.IDate,RentDetails.RentType,RentDetails.PackageType,RentDetails.Rate0,RentDetails.Rate1,RentDetails.Desc,RentDetails.Discount,RentDetails.Balance,RentDetails.NetAmount,ChartOfAccounts.AcName,ChartOfAccounts.Address1 ,PartyDetails.Address2,PartyDetails.Address3,PartyDetails.Mobile  FROM ((RentDetails  INNER JOIN ChartOfAccounts ON RentDetails.AccountID = ChartOfAccounts.AccountID ) INNER JOIN PartyDetails  ON RentDetails.AccountID=PartyDetails.AccountID) WHERE RentDetails.InvNo=" & Receiptprint1.Label3.Text & " "


            cmd = New OleDbCommand(sql, con)
            dReader = cmd.ExecuteReader
            While dReader.Read = True

                Receiptprint1.Label13.Text = dReader("IDate")
                Receiptprint1.Label10.Text = dReader("PackageType").ToString
                Receiptprint1.Label17.Text = dReader("Address1").ToString
                Receiptprint1.Label18.Text = dReader("Address2").ToString
                Receiptprint1.Label19.Text = dReader("Address3").ToString
                Receiptprint1.Label21.Text = dReader("Mobile").ToString
                TextBox10.Text = dReader("Balance").ToString
                Receiptprint1.Label6.Text = dReader("Desc").ToString
                TextBox1.Text = dReader("AccountID").ToString
                Receiptprint1.Label8.Text = dReader("Rate0").ToString
                Receiptprint1.Label11.Text = dReader("Rate1").ToString
                Receiptprint1.Label7.Text = dReader("RentType").ToString
                Receiptprint1.Label9.Text = dReader("Netamount").ToString
                Receiptprint1.Label25.Text = Val(dReader("Discount"))
                Receiptprint1.Label16.Text = dReader("AcName").ToString


            End While

        Catch ex As OleDbException

        Finally
            con.Close()

        End Try
    End Sub
End Class
