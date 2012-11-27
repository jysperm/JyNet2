<%@ WebHandler Language="VB" Class="Art_PV" %>

Imports System
Imports System.Web

Public Class Art_PV : Implements IHttpHandler
    
    Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
        If context.Request.QueryString("mode") = "post" Then
            Dim tArt As CArt = New CArt(context.Request.Form("id"))
            tArt.PV = tArt.PV + 1
            context.Response.Write(tArt.PV)
        End If
    End Sub
 
    Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class