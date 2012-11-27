<%@ WebHandler Language="VB" Class="Login" %>

Imports System
Imports System.Web

Public Class Login : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        context.Response.ContentType = "text/plain"
        If context.Request.QueryString("mode") = "post" Then
            Select Case context.Request.Form("do")
                Case "login"
                    'If UCase(Request.Form("vcode")) = Session("Vcode") Then
                    If UCase(context.Request.Form("vcode")) = "0" Then
                        '登录代码
                        Dim tr As Integer = CLogin.Login(context.Request.Form("user"), CString.MD5(context.Request.Form("pwd")))
                        If tr = 1 Then
                            context.Response.Write("OK")
                        Else
                            context.Response.Write("err")
                        End If
                    Else
                        context.Response.Write("vcodeerr")
                    End If
                Case "logout"
                    CLogin.LogOut()
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