Public Class CustomerLedger

    Private Sub CustomerLedger_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListView1.Columns.Add("SlNO", 75)
        ListView1.Columns.Add("Party Name", 200)
        ListView1.Columns.Add("Address", 200)
        ListView1.Columns.Add("Balance", 100)
        TOTL()
    End Sub

    Private Sub TOTL()
        Dim TempDbl As Double
        Dim total As Double



        For Each item1 In ListView1.Items
            If Double.TryParse(item1.SubItems.Item(3).Text, TempDbl) Then
                total += TempDbl
            End If
        Next





        Dim LI As New ListViewItem
        LI.Text = ""
        LI.SubItems.Add("")
        LI.SubItems.Add("TOTAL")




        LI.SubItems.Add(total)

        ' statements to add more subitems
        ListView1.Items.Add(LI)
        LI.Font = New Font("Verdana", 10, FontStyle.Bold)







    End Sub

    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click

        If CustomerLedgerFilter.ComboBox1.Text = "" Then
            CustomerLedgerFilter.ret1()
            TOTL()

        Else
            CustomerLedgerFilter.ret()
            TOTL()

        End If

    End Sub
End Class