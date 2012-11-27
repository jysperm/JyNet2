<%@ Page Language="VB" %>
<script runat="server">
    Public TU As CUser
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        If CLogin.Login() Then
            TU = New CUser(CLogin.GetUName)
        Else
            Response.Write("未登录或登录状态失效")
            Response.End()
        End If
    End Sub
</script>
<div style="text-align:left; width: 435px; height: 170px;">
UID:<%=TU.UID%>&nbsp;&nbsp;用户名：<%=TU.UName%><br />
昵称：<input id="infoNName" type="text" value="<%=TU.NName%>" /><br />
<img src="Images/Money.jpg"><span class="Money" style="float: none"><%=TU.Money%></span>
<span class="Exp" style="font-size: 15px; float: none;">积分：<span class="ExpIn"><%=TU.Exp%></span></span><br />
注册时间：<%=TU.RegTime%>&nbsp;&nbsp;注册IP：<%=TU.RegIP%><br />
QQ:<input id="infoQQ" type="text" value="<%=TU.QQ%>" /><br />
头像：<input type="button" value="点击这里设置头像" onclick="dialog.Close();SetLogo();" /><br />
邮箱：<input id="infoMail" type="text" value="<%=TU.Email.ToString%>" />
</div>