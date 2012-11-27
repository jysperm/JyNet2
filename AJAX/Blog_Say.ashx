<%@ WebHandler Language="VB" Class="Blog_Say" %>

Imports System
Imports System.Web
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Public Class Blog_Say : Implements IHttpHandler : Implements IRequiresSessionState
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        If context.Request.QueryString("mode") = "post" Then
            Select Case context.Request.Form("do")
                Case "getlogin"
                    If CLogin.Login Then
                        context.Response.Write(CLogin.GetUName.ToString)
                    Else
                        context.Response.Write("null")
                    End If
                Case "new"
                    If Not CType(CLogin.Login, Boolean) Then
                        If UCase(context.Request.Form("vcode")) <> context.Session("VCode") Then
                            context.Response.Write("请输入正确的验证码")
                            context.Session("VCode") = ""
                            context.Response.End()
                        Else
                            context.Session("VCode") = ""
                        End If
                    End If
                    Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                    Dim CMD As New OleDbCommand
                    CMD.Connection = CN
                    CMD.CommandText = "INSERT INTO [Blog_Say] (IsLogin, Body, UName,IP,AID) VALUES(@IsLogin, @Body, @UName,@IP,@AID)"
                    '向SQL语句中补充数据
                    CMD.Parameters.AddWithValue("@IsLogin", CLogin.Login)
                    Dim tBody As String = Replace(context.Request.Form("body"), "<", "&lt;")             '除去HTML标记
                    '替换表情
                    For i = 1 To 60
                        tBody = Replace(tBody, "[/表情" & i & "]", "<img src='http://" & CString.DoMain & "/Images/EM/" & i.ToString & ".gif' />")
                    Next
                    CMD.Parameters.AddWithValue("@Body", tBody)
                    CMD.Parameters.AddWithValue("@UName", IIf(CLogin.GetUName.ToString = "", "null", CLogin.GetUName.ToString))
                    CMD.Parameters.AddWithValue("@IP", context.Request.UserHostAddress.ToString)
                    CMD.Parameters.AddWithValue("@AID", context.Request.Form("aid"))
                    '执行命令
                    CMD.ExecuteNonQuery()
                    CMD.Dispose()
                    CN.Close()
                    context.Response.Write("OK")
                Case "get"
                    Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                    Dim CMD As New OleDbCommand
                    CMD.Connection = CN
                    CMD.CommandText = "SELECT * FROM [Blog_Say] WHERE [AID]=@AID ORDER BY [TheTime] DESC"
                    CMD.Parameters.AddWithValue("@AID", context.Request.Form("aid"))
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    Dim i As Integer = 0
                    Do While DR.Read And i < context.Request.Form("num")
                        context.Response.Write("<div class='SayItem'><span class='SayInfo'>By " & ShowUNameIP(DR("UName"), DR("IP"), DR("IsLogin")) & " At " & CString.MyTime(DR("TheTime")) & "</span>" & DR("Body") & "</div>")
                        i = i + 1
                    Loop
                    If i = 0 Then
                        context.Response.Write("还一个评论都没有呢！你来说说？")
                    End If
                    DR.Close()
                    CMD.Dispose()
                    CN.Close()
            End Select
        End If
    End Sub
    Public Function ShowUNameIP(ByVal UName As String, ByVal IP As String, ByVal IsLogin As String) As String
        If IsLogin Then
            ShowUNameIP = New CUser(UName).NName.ToString
        Else
            ShowUNameIP = IP
        End If
    End Function
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class