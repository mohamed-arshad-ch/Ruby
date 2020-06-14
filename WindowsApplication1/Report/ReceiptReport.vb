Public Class ReceiptReport

    Private Sub ReceiptReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListView1.Columns.Add("SlNO", 75)
        ListView1.Columns.Add("Invoice No", 90)
        ListView1.Columns.Add("Date", 100)
        ListView1.Columns.Add("Party Name", 200)
        ListView1.Columns.Add("Address", 200)
        ListView1.Columns.Add("Narration", 200)
        ListView1.Columns.Add("Discount", 100)
        ListView1.Columns.Add("Amount", 200)

    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems(0).SubItems(2).Text = "" Then

        Else
            Receipt.Show()
            Receipt.NumericUpDown1.Text = ListView1.SelectedItems(0).SubItems(1).Text


        End If
    End Sub

   

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class