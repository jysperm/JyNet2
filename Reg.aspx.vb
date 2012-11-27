
Partial Class Reg
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.QueryString("mode") = "post" Then
            Select Case Request.Form("do")
                Case "reg"
                    If UCase(Request.Form("vcode")) = Session("Vcode") Then
                        '注册代码
                        Response.Write(CUser.Reg(Request.Form("user"), Request.Form("pwd"), Request.Form("nname"), Request.Form("mail")))
                        Session("Vcode") = ""
                        Response.End()
                    Else
                        Response.Write("vcodeerr")
                    End If
                    Response.End()
                Case Else
                    Response.Write("ENF")
                    Response.End()
            End Select
        End If
    End Sub
End Class
