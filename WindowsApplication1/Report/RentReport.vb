Imports System.Data.OleDb
Public Class RentReport
    Dim con As OleDbConnection = myconn()
    Dim cmd As New OleDbCommand
    Dim da As New OleDbDataAdapter
    Dim dReader As OleDbDataReader
    Dim ds As New DataSet
    Dim dt As New DataTable
    Dim sql As String
    Dim agt As New ListView

    Private Sub RentReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListView1.Columns.Add("SlNO", 75)
        ListView1.Columns.Add("Invoice No", 100)
        ListView1.Columns.Add("Date", 100)
        ListView1.Columns.Add("Party Name", 200)
        ListView1.Columns.Add("Address", 200)
        ListView1.Columns.Add("Rent ", 150)
        ListView1.Columns.Add("Package", 150)
        ListView1.Columns.Add("Discount ", 150)
        ListView1.Columns.Add("Netamount ", 150)




        TOTL()

    End Sub


   



    Private Sub RefreshToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RefreshToolStripMenuItem.Click
        If RentReportFilter.ComboBox1.Text = "" Or RentReportFilter.ComboBox2.Text = "" Then
            RentReportFilter.ret1()
            TOTL()
        Else
            RentReportFilter.ret()
            TOTL()
        End If
       

    End Sub

    Private Sub TOTL()
        Dim TempDbl As Double
        Dim total As Double
        Dim TempDb2 As Double
        Dim total1 As Double


        For Each item1 In ListView1.Items
            If Double.TryParse(item1.SubItems.Item(8).Text, TempDbl) Then
                total += TempDbl
            End If
        Next

        For Each item1 In ListView1.Items
            If Double.TryParse(item1.SubItems.Item(7).Text, TempDb2) Then
                total1 += TempDb2
            End If
        Next



        Dim LI As New ListViewItem
        LI.Text = ""
        LI.SubItems.Add("")
        LI.SubItems.Add("")
        LI.SubItems.Add("")
        LI.SubItems.Add("")
        LI.SubItems.Add("TOTAL")
        LI.SubItems.Add("")
        LI.SubItems.Add(total1)
        LI.SubItems.Add(total)

        ' statements to add more subitems
        ListView1.Items.Add(LI)
        LI.Font = New Font("Verdana", 10, FontStyle.Bold)







    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems(0).SubItems(2).Text = "" Then

        Else
            Form1.Show()
            Form1.NumericUpDown1.Text = ListView1.SelectedItems(0).SubItems(1).Text


        End If
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel2.Paint

    End Sub
End Class