<%@ WebHandler Language="VB" Class="SetLogo" %>

Imports System
Imports System.Web

Public Class SetLogo : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        If context.Request.QueryString("mode") = "post" Then
            Select Case context.Request.Form("do")
                Case "setlogo"
                    If Len(context.Request.Form("url")) <> 0 Then
                        If context.Request.Form("url").Length < 200 And InStr(context.Request.Form("url"), "<") < 1 And InStr(context.Request.Form("url"), ">") < 1 And InStr(context.Request.Form("url"), Chr(34)) < 1 And InStr(context.Request.Form("url"), Chr(39)) < 1 Then
                            If CLogin.Login() Then
                                Dim TU = New CUser(CLogin.GetUName)
                                TU.Logo = context.Request.Form("url")
                                context.Response.Write("OK")
                            Else
                                context.Response.Write("未登录或登录状态失效")
                            End If
                        Else
                            context.Response.Write("非法网址")
                        End If
                    Else
                        context.Response.Write("不能设置为空")
                    End If
                Case "setinfo"
                    If CLogin.Login() Then
                        Dim TU = New CUser(CLogin.GetUName)
                        TU.NName = context.Request.Form("nname")
                        TU.QQ = context.Request.Form("qq")
                        TU.EMail = context.Request.Form("mail")
                        context.Response.Write("OK")
                    Else
                        context.Response.Write("未登录或登录状态失效")
                    End If
                Case "setpwd"
                    If CLogin.Login() Then
                        If CLogin.Login(CLogin.GetUName, CString.MD5(context.Request.Form("old"))) Then
                            Dim TU = New CUser(CLogin.GetUName)
                            TU.PWD = CString.MD5(context.Request.Form("new"))
                            CLogin.LogOut()
                            context.Response.Write("OK")
                        Else
                            context.Response.Write("原密码错误")
                        End If
                    Else
                        context.Response.Write("未登录或登录状态失效")
                    End If
                Case Else
                    context.Response.Write("调用错误")
                    
            End Select
        End If
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class