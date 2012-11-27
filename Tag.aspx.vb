Imports System.Data.OleDb

Partial Class _Tag
    Inherits System.Web.UI.Page
    Public I As Integer
    Public sNum, sOrder, sTag As String
    Public CN As OleDbConnection = CConnDB.ConnToDB(0)
    Public CMD As New OleDbCommand()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CMD.Connection = CN

    End Sub
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        CMD.Dispose()
        CN.Close()
    End Sub
End Class
