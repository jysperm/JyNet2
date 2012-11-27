Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Public Class CConnDB
    Public Shared Function ConnToDB(Optional ByVal DBID As Integer = 0) As OleDbConnection
        ConnToDB = New OleDbConnection
        ConnToDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "../App_Data" & "/" & CFile.XMLKey(CFile.LoadFile("../Admin/DBID.ini"), DBID.ToString) & ";Persist Security Info=False"
        ConnToDB.Open()
    End Function
End Class