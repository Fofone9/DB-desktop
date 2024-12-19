Imports System.Data.OleDb
Module Module1
    Public con As New OleDbConnection
    Sub main()
        con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\mrpaa\OneDrive\Desktop\oledb\bd.mdb"
        Dim frmCustomer As New frmCustomer
        frmCustomer.ShowDialog()
    End Sub
End Module
