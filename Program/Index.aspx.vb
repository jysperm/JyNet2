Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Partial Class Program_Index
    Inherits System.Web.UI.Page
    Public TypeID As Integer = 2
    Public I As Integer
    Public CN As OleDbConnection = CConnDB.ConnToDB(0)
    Public CMD As New OleDbCommand()
    Public CMDN As New OleDbCommand()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CMD.Connection = CN
        CMDN.Connection = CN
    End Sub
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        CMD.Dispose()
        CMDN.Dispose()
        CN.Close()
    End Sub
End Class
