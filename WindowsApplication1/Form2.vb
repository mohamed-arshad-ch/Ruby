Public Class Form2

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.IsMdiContainer = True
    End Sub

    Private Sub RentDetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RentDetailsToolStripMenuItem.Click
        Form1.MdiParent = Me
        Form1.Show()
    End Sub

    Private Sub RentDetailsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RentDetailsToolStripMenuItem1.Click
        RentReportFilter.MdiParent = Me
        RentReportFilter.Show()
    End Sub

    Private Sub CoustomerListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CoustomerListToolStripMenuItem.Click
        CustomerLedgerFilter.MdiParent = Me
        CustomerLedgerFilter.Show()
    End Sub

    Private Sub ReceiptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptToolStripMenuItem.Click
        Receipt.MdiParent = Me
        Receipt.Show()
    End Sub

    Private Sub ReceiptToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReceiptToolStripMenuItem1.Click
        RecieptFilter.MdiParent = Me
        RecieptFilter.Show()
    End Sub
End Class