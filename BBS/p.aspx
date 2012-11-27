<%@ Page Language="VB" AutoEventWireup="false" CodeFile="p.aspx.vb" Inherits="BBS_p" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>
        <%= GetP("Title").ToString%>_精英盒子.Net 论坛</title>
    <link href="../CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/BBS.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebAjax.js" type="text/javascript"></script>
    <script src="../JS/WebDialog.js" type="text/javascript"></script>
    <script src="../JS/WebCommon.js" type="text/javascript"></script>
    <script src="../JS/KE.js" type="text/javascript"></script>
    <script src="../JS/User.js" type="text/javascript"></script>
    <script type="text/javascript">
        function ILoad() {
            window.location = window.location;
        }
        function OnNewP() {
            if("OK"!=executeHttpRequest("POST", "../AJAX/BBS.ashx?mode=post","do=getlogin")) {
                alert("你还没登陆呢，登录再发帖吧");
                Login(ILoad); 
                return;
            }
            var r = executeHttpRequest("POST", "../AJAX/BBS.ashx?mode=post", "do=r&pid=<%= GetP("ID").ToString%>&body=" + encodeURIComponent($("RIn").value));
            if(r=="OK")
            {
                alert("回复成功");
                ILoad();
            }
            else
            {
                alert("回复过程中发送错误");
            }
        }
    </script>
</head>
<body>
    <script type="text/javascript">
        var dialog = new WebDialog();
        dialog.ImgZIndex = 107;
        dialog.DialogZIndex = 108;
        dialog.Width = 400;
        KE.init({
            id: 'RIn'
        });
    </script>
    <form id="FormMain" runat="server">
    <div id="MainFrame">
        <div id="HeadBar">
            加载中...
        </div>
        <div class="HeadArea">
            <div class="HeadLogo">
                <img alt="Logo" src="../Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable MainTablep">
            <div id="PTitle">
                <div id="TitleIn">
                    <%= GetP("Title").ToString%></div>
                <div id="TitleR">
                    <%= New CBBSRe(GetP("ID")).GetReNum()%>回复/<%= GetP("PV").ToString%>浏览</div>
            </div>
            <div id="tDir">
                *<a href="http://jybox.net">精英盒子.Net</a> -> <a href="index.aspx">论坛</a> -> <a href="index.aspx?id=<%= GetB("ID").ToString%>">
                    <%= GetB("BName").ToString%></a> ->
                <%= GetP("Title").ToString%>
            </div>
            <div class="PItem">
                <div class="PIL">
                    <img alt="<%= New CUser(GetP("UName").ToString).NName%>" src="<%= New CUser(GetP("UName").ToString).Logo.ToString%>"
                        class="ULogo" />
                    <br />
                    <span class="NName">
                        <%= New CUser(GetP("UName").ToString).NName%></span><br />
                    (<%= New CUser(GetP("UName").ToString).UName%>)<br />
                    积分<span class="EXP"><%= New CUser(GetP("UName").ToString).Exp.ToString%></span><br />
                    <img src="../Images/Money.jpg" alt="金币" /><%= New CUser(GetP("UName").ToString).Money.ToString%><br />
                </div>
                <div class="PIR">
                    <div class="BIN">
                        <%= GetP("Body").ToString%>
                    </div>
                </div>
                <div class="Clear">
                </div>
            </div>
            <%
            
                CMD.CommandText = "SELECT * FROM [BBS_R] WHERE [PID]=" & TID & " ORDER BY [tTime]"
                Dim DR As OleDbDataReader = CMD.ExecuteReader
                Dim F = (ListPage - 1) * 30 + 1
                Dim T = (ListPage - 1) * 30 + 1 + 30
                Dim I = 1
                Do While DR.Read And I < T
                    If I >= F And I < T Then
            %>

            <div class="PItem">
                <div class="PIL">
                    <img alt="<%= New CUser(DR("UName").ToString).NName%>" src="<%= New CUser(DR("UName").ToString).Logo.ToString%>"
                        class="ULogo" />
                    <br />
                    <span class="NName">
                        <%= New CUser(DR("UName").ToString).NName%></span><br />
                    (<%= New CUser(DR("UName").ToString).UName%>)<br />
                    积分<span class="EXP"><%= New CUser(DR("UName").ToString).Exp.ToString%></span><br />
                    <img src="../Images/Money.jpg" alt="金币" /><%= New CUser(DR("UName").ToString).Money.ToString%><br />
                </div>
                <div class="PIR">
                    <div class="BIN">
                        <%= DR("Body").ToString%>
                    </div>
            </div>
            <div class="Clear">
            </div>
        </div>
        <%
        End If
        I = I + 1
    Loop
    DR.Close()
        %>
        <div id="NewR">
        <%If CLogin.Login Then%>
        [<%= New CUser(CLogin.GetUName).NName%>]请在这里回复帖子：<br />

        <textarea id="RIn" name="context" cols="100"  rows="8"></textarea>
        <script type="text/javascript">
            setTimeout("KE.create('RIn');", "100");
        </script>
        <input id="NewROK" type="button" value="发表" onclick="OnNewP();" />
        <div class="Clear"></div>
        <%Else%>
        您还没登录呢！不能回帖子
        <input type="button" value="登录" onclick="Login(ILoad);" /><input type="button" value="注册"
                            onclick="window.location = 'http://jybox.net/Reg.aspx'" />

        <%End If%>
        </div>
    </div>
    </div>
    
    <div id="Footer">
    </div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    
    <script type="text/javascript" src="../JS/Frame.js"></script>
</body>
</html>
