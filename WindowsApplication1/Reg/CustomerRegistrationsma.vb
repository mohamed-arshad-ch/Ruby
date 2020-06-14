Public Class CustomerRegistrationsma

    Private Sub CustomerRegistrationsma_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        loadaccc()
    End Sub

    Private Sub inseret()
        Dim today As String = DateTime.Now.ToString("dd/MM/yy")
        issucess = insert("insert into ChartOfAccounts values ('" & TextBox5.Text & "','100','" & TextBox1.Text & "','" & TextBox2.Text & "','" & today & "','-1','" & TextBox1.Text & "')")
        issucess = insert("insert into PartyDetails values ('" & TextBox5.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','','','')")
    End Sub

    Private Sub loadaccc()
        loadaccount("select top 1 AccountID from ChartOfAccounts order by AccountID desc", "AccountID")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("Pls Enter Valid Data!", "ALERT", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Else
            inseret()
            Me.Close()
        End If
    End Sub
End Class