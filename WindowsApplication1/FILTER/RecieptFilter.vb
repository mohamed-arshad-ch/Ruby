Public Class RecieptFilter

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Then
           
            FIL()
            ReceiptReport.Show()


        Else
            FIL1()
            ReceiptReport.Show()
        End If
    End Sub

    Private Sub FIL()
        findthis("SELECT ReceiptPayment.RPNo, ReceiptPayment.TDate,ChartOfAccounts.AcName,ChartOfAccounts.Address1 ,ReceiptPayment.Narration,ReceiptPayment.Dis,ReceiptPayment.Amount FROM (ReceiptPayment  INNER JOIN ChartOfAccounts ON ReceiptPayment.AccountID = ChartOfAccounts.AccountID ) WHERE ReceiptPayment.TypeCode=15   ORDER BY ReceiptPayment.RPNo ")
        LoadReceiptR(ReceiptReport.ListView1)
    End Sub
    Private Sub FIL1()
        findthis("SELECT ReceiptPayment.RPNo, ReceiptPayment.TDate,ChartOfAccounts.AcName,ChartOfAccounts.Address1 ,ReceiptPayment.Narration,ReceiptPayment.Dis,ReceiptPayment.Amount FROM (ReceiptPayment  INNER JOIN ChartOfAccounts ON ReceiptPayment.AccountID = ChartOfAccounts.AccountID ) WHERE ChartOfAccounts.AcName='" & ComboBox1.Text & "' AND ReceiptPayment.TypeCode=15   ORDER BY ReceiptPayment.RPNo ")
        LoadReceiptR(ReceiptReport.ListView1)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()

    End Sub

    Private Sub RecieptFilter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        combofill(ComboBox1, "select AcName from ChartOfAccounts where GroupID=100", "AcName")
    End Sub
End Class