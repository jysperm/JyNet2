<%@ Page Language="VB" AutoEventWireup="false" CodeFile="My.aspx.vb" Inherits="_My" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>个人中心_精英盒子.Net</title>
    <link href="CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="CSS/My.css" rel="stylesheet" type="text/css" />
    <link href="CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="JS/WebAjax.js" type="text/javascript"></script>
    <script src="JS/WebDialog.js" type="text/javascript"></script>
    <script src="JS/WebCommon.js" type="text/javascript"></script>
<script src="../JS/User.js" type="text/javascript"></script>
    <%  If IsLoged Then%>
    <script type="text/javascript">
    function SetLogo()
    {
        dialog.Width = 470;
        dialog.Text='设置头像';
        dialog.Content =executeHttpRequest("GET", "UI/SetLogo.htm",null);
        dialog.Show(1);
        dialog.OK = function()
        {
            if(!($("OpFromWeb").checked || $("OpFromPC").checked))
            {
                alert("请选择一种头像类型(网络图片或本机上传)");
                return;
            }
            if($("OpFromWeb").checked)
            {
                var result = executeHttpRequest("POST", "AJAX/My.ashx?mode=post","do=setlogo&url=" + encodeURIComponent($("TextFromWeb").value));
                if(result=="OK")
                {
                    alert("设置成功");
                    window.location="My.aspx";
                }
                else
                {
                    alert(result);
                }
            }
            else
            {
                var ifr_window = window.frames["uplogo"];
                ifr_window.FromPC.submit();
            }
        }
    }
    function SetInfo()
    {
        dialog.Width = 470;
        dialog.Text='设置/查看个人信息';
        dialog.Content =executeHttpRequest("GET", "UI/SetInfo.aspx",null);
        dialog.Show(1);
        dialog.OK = function()
        {
            var RV = /^[^<>,\n]{1,16}$/;
            if(!RV.test($("infoNName").value))
            {
                alert("昵称格式错误\n昵称要求1-16位字符(允许中文)，不允许\"<\"、\">\"、\",\"等特殊字符");
                return;
            }
            RV = /^[^<>,\n]{0,12}$/;
            if(!RV.test($("infoQQ").value))
            {
                alert("QQ格式错误\nQQ不能长于12位");
                return;
            }
            RV = /^\w+([-+.]\w+)*@\w+([-.]\\w+)*\.\w+([-.]\w+)*$/;
            if(!RV.test($("infoMail").value))
            {
                alert("邮箱格式错误\n请输入标准的邮箱地址");
                return;
            }
            var result = executeHttpRequest("POST", "AJAX/My.ashx?mode=post","do=setinfo&nname=" + encodeURIComponent($("infoNName").value) + "&qq=" + encodeURIComponent($("infoQQ").value) + "&mail=" + encodeURIComponent($("infoMail").value));
            if(result=="OK")
            {
                $("NName").innerHTML=$("infoNName").value;
                alert("设置成功");
                dialog.Close();
            }
            else
            {
                alert(result);
            }
        }
    }
    function SetPWD()
    {
        dialog.Width = 310;
        dialog.Text='修改密码';
        dialog.Content =executeHttpRequest("GET", "UI/SetPWD.htm",null);
        dialog.Show(1);
        dialog.OK = function()
        {
            var RV = /^\w{6,16}$/;
            if(!RV.test($("NewPWD").value) || $("NewPWD").value != $("RePWD").value)
            {
                alert("密码格式错误\n密码要求6-16位字符，两次输入的密码必须相同");
                return;
            }
            var result = executeHttpRequest("POST", "AJAX/My.ashx?mode=post","do=setpwd&old=" + encodeURIComponent($("OldPWD").value) + "&new=" + encodeURIComponent($("NewPWD").value));
            if(result=="OK")
            {
                alert("密码修改成功，请重新登录");
                window.location="My.aspx";
            }
            else
            {
                alert(result);
            }
        }
    }
    function DoLogOut()
    {
        var result = executeHttpRequest("POST", "AJAX/Login.ashx?mode=post","do=logout");
        window.location="My.aspx";
    }
    function SelectOp()
    {
        $("OpFromWeb").checked=true;
    }
    function EndUp()
    {
        var ifr_window = window.frames["uplogo"];
        if(ifr_window.re.innerHTML=="OK")
        {
            var result = executeHttpRequest("POST", "AJAX/My.ashx?mode=post","do=setlogo&url=" + encodeURIComponent(ifr_window.Url.innerHTML));
            if(result=="OK")
            {
                alert("设置成功");
                window.location="../My.aspx";
            }
            else
            {
                alert(result);
            }
        }
        else
        {
            alert(ifr_window.re.innerHTML);
        }
    }
    </script>
    <%Else%>
    <script type="text/javascript">
    function OKClick()
    {
    var RV = /^[^<>,\n]{4,16}$/
    if(!RV.test($("TextU").value))
    {
        dialog.Text="提示-用户名格式错误";
        dialog.Content ="<div style='color:red;text-align:left;'>请输入正确的用户名</div>";
        dialog.Show(1);
        return;
    }
    RV = /^\w{6,16}$/
    if(!RV.test($("TextP").value))
    {
        dialog.Text="提示-密码格式错误";
        dialog.Content ="<div style='color:red;text-align:left;'>请输入正确的密码</div>";
        dialog.Show(1);
        return;
    }
    var result = executeHttpRequest("POST", "AJAX/Login.ashx?mode=post","do=login&user=" + encodeURIComponent(document.getElementById("TextU").value) + "&pwd=" + encodeURIComponent(document.getElementById("TextP").value) + "&vcode=0")
    if(result=="OK")
    {
        window.location.reload();
        return;
    }
    else if(result=="err")
    {
        dialog.Text="提示-用户名或密码错误";
        dialog.Content ="<div style='color:red;text-align:left;'>用户名或密码错误，请重新登录</div>";
        dialog.OK = function()
        {
            dialog.Close();
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
            dialog.Close();
        }
        dialog.Show(1);
        return;
    }
    }
    </script>
    <%End If%>
</head>
<body>
    <script type="text/javascript">
        var dialog = new WebDialog();
        dialog.ImgZIndex = 107;
        dialog.DialogZIndex = 108;
        dialog.Width = 400;
    </script>
    <form id="FormMain" runat="server">
    <div id="MainFrame">
        <div id="HeadBar">
            加载中
        </div>
        <div class="HeadArea">
            <div class="HeadLogo">
                <img alt="Logo" src="Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable">
            <%  If IsLoged Then%>
            <div id="UInfo">
            <div id="UIn">
            <img id="Logo" alt="用户头像" src="<%=TU.Logo.ToString%>" />
            <span class="NName" id="NName"><%=TU.NName.ToString%></span><br />
            <img class="ImgMoney" src="Images/Money.jpg" /><span class="Money"><%=TU.Money.ToString%></span><br />
            <span class="Exp">积分：<span class="ExpIn"><%=TU.Exp.ToString%></span></span>
            <div class="Clear"></div>
            <input id="Pay" type="button" value="充值" />
            <input id="ToIndex" type="button" value="首页" />
            <input id="LogOut" type="button" value="退出" onclick="DoLogOut();" />
            </div>
            </div>
            <div id="DivTop">2</div>
            <div id="NaBar">
            <div id="exNav">
                <ul>
                    <li><a href="My.aspx">个人中心主页</a></li>
                    <li><a href="#" onclick="javascript:SetLogo();">修改头像</a></li>
                    <li><a href="#" onclick="javascript:SetInfo();">修改/查看个人信息</a></li>
                    <li><a href="#" onclick="javascript:SetPWD();">修改密码</a></li>
                    <li><a href="#" onclick="javascript:Money();">金币/积分管理</a></li>
                    <li><a href="#" onclick="javascript:MyTools();">我的应用</a></li>
                    <li><a href="#" onclick="javascript:Mail();">站内信件</a></li>
                    <li><a href="#" onclick="javascript:MyDocs();">我的投稿</a></li>
                    <li><a href="#" onclick="javascript:BBS();">我的帖子</a></li>
                    <li><a href="#" onclick="javascript:MyFile();">我的文件</a></li>
                    <li><a href="#" onclick="javascript:Setting();">设置</a></li>
                </ul>
            </div></div>
            <div id="DivAjax">4</div>
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            <%Else%>
            <div id="MyLogin">
                <h2>=请登录个人中心=</h2>
                <input id="TextU" type="text" tabindex="1" />
                用户名：<br />
                <input id="TextP" type="password" tabindex="2" />
                密码：
                <div class="Clear">
                </div>
                <input id="OK" type="button" value="登录" onclick="javascript:OKClick();" 
                    tabindex="3" />
            </div>
            <%End If%>
        </div>
        
    </div>
    <div class="Clear"></div>
    <div id="Footer"></div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="JS/Frame.js"></script>
</body>
</html>
