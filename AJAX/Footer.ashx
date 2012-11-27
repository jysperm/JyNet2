<%@ WebHandler Language="VB" Class="Footer" %>

Imports System
Imports System.Web

Public Class Footer : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        If context.Request.QueryString("mode") = "post" Then
            Select Case context.Request.Form("do")
                Case "get"
                    context.Response.Write(Replace(CFile.LoadFile(context.Server.MapPath("~/UI/Footer.htm")), "[Ver]", CFile.XMLKey(CFile.LoadFile(context.Server.MapPath("~/Admin/Setting.ini")), "Ver")))
                Case Else
                    context.Response.Write("页脚加载错误")
            End Select
        End If
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class