Imports System.Data.OleDb
Partial Class Index
    Inherits System.Web.UI.Page
    Public TU As CUser, IsLoged As Integer = 0
    Public I As Integer
    Public CN As OleDbConnection = CConnDB.ConnToDB(0)
    Public CMD As New OleDbCommand()
    Public CMDN As New OleDbCommand()
    Public CMDB As New OleDbCommand()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CMD.Connection = CN
        CMDN.Connection = CN
        CMDB.Connection = CN
        IsLoged = CLogin.Login()
        TU = New CUser(CLogin.GetUName)
        '----------------------------
    End Sub

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        CMD.Dispose()
        CMDN.Dispose()
        CMDB.Dispose()
        CN.Close()
    End Sub
End Class
