<%@ WebHandler Language="VB" Class="BBS" %>

Imports System
Imports System.Web
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Public Class BBS : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "text/plain"
        If context.Request.QueryString("mode") = "post" Then
            Select Case context.Request.Form("do")
                Case "getlogin"
                    If CLogin.Login Then
                        context.Response.Write("OK")
                    Else
                        context.Response.Write("未登录")
                    End If
                Case "new"
                    If CLogin.Login Then
                        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                        Dim CMD As New OleDbCommand
                        CMD.Connection = CN
                        CMD.CommandText = "INSERT INTO [BBS_P] (BID, UName, tTime,LastTime,IP,Body,Title,Link,UpSet,PV) VALUES(@BID, @UName, @tTime,@LastTime,@IP,@Body,@Title,@Link,@UpSet,@PV)"
                        '向SQL语句中补充数据
                        CMD.Parameters.AddWithValue("@BID", Int(context.Request.Form("bid").ToString))
                        CMD.Parameters.AddWithValue("@UName", CLogin.GetUName.ToString)
                        CMD.Parameters.AddWithValue("@tTime", Now.ToString)
                        CMD.Parameters.AddWithValue("@LastTime", Now.ToString)
                        CMD.Parameters.AddWithValue("@IP", context.Request.UserHostAddress.ToString)
                        CMD.Parameters.AddWithValue("@Body", context.Request.Form("body").ToString)
                        Dim tTitle As String = Replace(context.Request.Form("title"), "<", "&lt;")
                        CMD.Parameters.AddWithValue("@Title", tTitle)
                        CMD.Parameters.AddWithValue("@Link", "")
                        CMD.Parameters.AddWithValue("@UpSet", 0)
                        CMD.Parameters.AddWithValue("@PV", 0)
                        CMD.ExecuteNonQuery()
                        CMD.Dispose()
                        CN.Close()
                        context.Response.Write("OK")
                    Else
                        context.Response.Write("未登录或登录状态失效")
                    End If
                Case "r"
                    If CLogin.Login Then
                        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                        Dim CMD As New OleDbCommand
                        CMD.Connection = CN
                        CMD.CommandText = "INSERT INTO [BBS_R] (PID, UName, tTime,IP,Body) VALUES(@PID, @UName, @tTime,@IP,@Body)"
                        '向SQL语句中补充数据
                        CMD.Parameters.AddWithValue("@PID", Int(context.Request.Form("pid").ToString))
                        CMD.Parameters.AddWithValue("@UName", CLogin.GetUName.ToString)
                        CMD.Parameters.AddWithValue("@tTime", Now.ToString)
                        CMD.Parameters.AddWithValue("@IP", context.Request.UserHostAddress.ToString)
                        CMD.Parameters.AddWithValue("@Body", context.Request.Form("body").ToString)
                        CMD.ExecuteNonQuery()
                        CMD.Dispose()
                        CN.Close()
                        '如果ID为1，增加1金币
                        If context.Request.Form("pid").ToString = "1" Then
                            Dim TU As New CUser(CLogin.GetUName)
                            TU.Money = TU.Money + 1
                        End If
                        context.Response.Write("OK")
                    Else
                        context.Response.Write("未登录或登录状态失效")
                    End If
                Case Else
                    context.Response.Write("ENF")
            End Select
        End If
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class