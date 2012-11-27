<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Reg.aspx.vb" Inherits="Reg" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>注册帐号_精英盒子.Net</title>
    <link href="CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="CSS/Reg.css" rel="stylesheet" type="text/css" />
    <link href="CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="JS/WebAjax.js" type="text/javascript"></script>
    <script src="JS/WebDialog.js" type="text/javascript"></script>
    <script src="JS/WebCommon.js" type="text/javascript"></script>
    <script src="JS/User.js" type="text/javascript"></script>
    <script type="text/javascript">
        function ReVCode()
        {
            times++;
            document.getElementById("VCodeImg").src="";
            document.getElementById("VCodeImg").src="UI/VCode.aspx?times=" + times ;
        }
        function OKClick()
        {
            if(!$("ChB_Eula").checked)
            {
                
                dialog.Text="提示-你没有同意用户协议&声明";
                dialog.Content ="<div style='color:red;text-align:left;'>请您认真阅读并同意<a href='Eula.aspx'>用户声明&协议</a></div>";
                dialog.Show(1);
                return;
            }
            var RV = /^[^<>,\n]{4,16}$/
            if(!RV.test($("Text_UName").value))
            {
                dialog.Text="提示-用户名格式错误";
                dialog.Content ="<div style='color:red;text-align:left;'>用户名要求4-16位字符(允许中文)，不允许\"<\"、\">\"、\",\"等特殊字符<br /><br />正则表达式(高级):/^[^<>,\\n]{4,16}$/</div>";
                dialog.Show(1);
                return;
            }
            RV = /^[^<>,\n]{1,16}$/
            if(!RV.test($("Text_NName").value))
            {
                dialog.Text="提示-昵称格式错误";
                dialog.Content ="<div style='color:red;text-align:left;'>昵称要求1-16位字符(允许中文)，不允许\"<\"、\">\"、\",\"等特殊字符<br /><br />正则表达式(高级):/^[^<>,\\n]{1,16}$/</div>";
                dialog.Show(1);
                return;
            }
            RV = /^\w{6,16}$/
            if(!RV.test($("PWD").value) || $("PWD").value != $("RePWD").value)
            {
                dialog.Text="提示-密码格式错误";
                dialog.Content ="<div style='color:red;text-align:left;'>密码要求6-16位字符，两次输入的密码必须相同<br /><br />正则表达式(高级):/^\\w{6,16}$/</div>";
                dialog.Show(1);
                return;
            }
            RV = /^\w+([-+.]\w+)*@\w+([-.]\\w+)*\.\w+([-.]\w+)*$/
            if(!RV.test($("Text_Mail").value))
            {
                dialog.Text="提示-邮箱格式错误";
                dialog.Content ="<div style='color:red;text-align:left;'>请输入标准的邮箱地址<br /><br />正则表达式(高级):/^\\w+([-+.]\\w+)*@\\w+([-.]\\\\w+)*\\.\\w+([-.]\\w+)*$/</div>";
                dialog.Show(1);
                return;
            }
            RV = /^\w{4}$/
            if(!RV.test($("VCode").value))
            {
                dialog.Text="提示-验证码格式错误";
                dialog.Content ="<div style='color:red;text-align:left;'>请输入4位验证码(均为字母，不区分大小写)<br /><br />正则表达式(高级):/^\\w{4}$/</div>";
                dialog.Show(1);
                return;
            }
            var result = executeHttpRequest("POST", "Reg.aspx?mode=post","do=reg&user=" + encodeURIComponent(document.getElementById("Text_UName").value) + "&pwd=" + encodeURIComponent(document.getElementById("PWD").value) + "&nname=" + encodeURIComponent(document.getElementById("Text_NName").value) + "&mail=" + encodeURIComponent(document.getElementById("Text_Mail").value) + "&vcode=" + document.getElementById("VCode").value)
            if(result=="sameuser")
            {
                dialog.Text="提示-用户名重复";
                dialog.Content ="<div style='color:red;text-align:left;'>用户名重复，请重新输入一个用户名</div>";
                dialog.OK = function()
                {
                    ReVCode();
                    dialog.Close();
                }
                dialog.Show(1);
                return;
            }
            else if(result=="OK")
            {
                dialog.Text="注册成功！";
                dialog.Content ="<div style='color:red;text-align:left;'>注册成功！点击确定跳转到首页</div>";
                dialog.OK = function()
                {
                    window.location="Index.aspx";
                }
                dialog.Show(1);
                return;
            }
            else if(result=="vcodeerr")
            {
                dialog.Text="提示-验证码错误";
                dialog.Content ="<div style='color:red;text-align:left;'>验证码错误，请重新输入验证码</div>";
                dialog.OK = function()
                {
                    ReVCode();
                    dialog.Close();
                }
                dialog.Show(1);
                return;
            }
            else
            {
                dialog.Text="提示-不明错误";
                dialog.Content ="<div style='color:red;text-align:left;'>遇到不明错误。如频繁出现此错误请联系管理员(m@jybox.net)</div>";
                dialog.OK = function()
                {
                    ReVCode();
                    dialog.Close();
                }
                dialog.Show(1);
                return;
            }
        }
    </script>
</head>
<body>
    <script type="text/javascript">
        var times=0;
        var dialog = new WebDialog();
        dialog.ImgZIndex = 107;
        dialog.DialogZIndex = 108;
        dialog.Width = 400;
    </script>
    <form id="FormMain" runat="server">
    <div id="MainFrame">
        <div id="HeadBar">
        加载中...
        </div>
        <div class="HeadArea">
            <div class="HeadLogo">
                <img alt="Logo" src="Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable">
        <div id="RegMain">
            <br />
            <div class="FRight">
            <input id="Text_UName" type="text" tabindex="1" /><br />
            <input id="Text_NName" type="text" tabindex="2" /><br />
            <input id="PWD" type="password" tabindex="3" /><br />
            <input id="RePWD" type="password" tabindex="4" /><br />
            <input id="Text_Mail" type="text" tabindex="5" /><br />
            <a href="#" onclick="javascript:ReVCode();"><img id="VCodeImg" alt="点击更换验证码" src="UI/VCode.aspx" /></a>
            <input id="VCode" type="text" tabindex="6" />
            </div>
            用户名：<br />
            昵称：<br />
            密码：<br />
            重复密码：<br />
            邮箱：<br />
            验证码：<br />
            <div class="Clear"></div>
            <input id="OK" type="button" value="OK" onclick="javascript:OKClick();" 
                tabindex="8" />
            <input id="ChB_Eula" type="checkbox" tabindex="7" />同意<a target="_blank" href="Eula.aspx">用户协议&amp;声明</a><br />
        </div>
        </div>
    </div>
    <div id="Footer"></div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="JS/Frame.js"></script>
</body>
</html>
