Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.Web.HttpContext
Public Class CConnDB
    Public Shared Function ConnToDB(Optional ByVal DBID As Integer = 0) As OleDbConnection
        ConnToDB = New OleDbConnection
        ConnToDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Current.Server.MapPath("~/App_Data") & "/" & CFile.XMLKey(CFile.LoadFile(Current.Server.MapPath("~/Admin/DBID.ini")), DBID.ToString) & ";Persist Security Info=False"
        ConnToDB.Open()
    End Function
End Class