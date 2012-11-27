<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Index.aspx.vb" Inherits="BBS_Index" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>
        <%= GetTitle()%>精英盒子论坛_精英盒子.Net</title>
    <meta name="keywords" content="论坛,交流,讨论,版块,帖子,JyNet,精英王子,精英盒子,精英盒子.Net,开源" />
    <meta name="description" content="这里是我们自行研发的供用户交流的论坛，欢迎讨论编程、计算机应用、灌水" />
    <link href="../CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebAjax.js" type="text/javascript"></script>
    <script src="../JS/WebDialog.js" type="text/javascript"></script>
    <script src="../JS/WebCommon.js" type="text/javascript"></script>
    <link href="../CSS/BBS.css" rel="stylesheet" type="text/css" />
    <script src="../JS/User.js" type="text/javascript"></script>
    <script src="../JS/KE.js" type="text/javascript"></script>
    <%If TID <> -1 Then%>
    <script type="text/javascript" defer="defer">
        function ALoad() {
            window.location = window.location;
        }
        function PNewP() {
            if("OK"!=executeHttpRequest("POST", "../AJAX/BBS.ashx?mode=post","do=getlogin")) {
                alert("你还没登陆呢，登录再发帖吧");
                Login(ALoad); 
                return;
            }
            dialog.Width = 720;
            dialog.Text = '发表新帖子';
            KE.init({
                id: 'In'
            });
            dialog.Content = "标题：<input id='Title' type='text' style='width: 649px;' /><br /><textarea id='In' name='context' cols='100' rows='8' style='width: 700px; height: 300px;'></textarea>"
            setTimeout("KE.create('In');", "100");
            dialog.Show(1);
            dialog.OK = function () {
                var r = executeHttpRequest("POST", "../AJAX/BBS.ashx?mode=post", "do=new&bid=<%= GetB("ID").ToString%>&title=" + encodeURIComponent($("Title").value) + "&body=" + encodeURIComponent($("In").value));
                if(r=="OK")
                {
                    alert("发帖成功");
                    ALoad();
                }
                else
                {
                    alert("发帖过程中发送错误");
                }
            }
        }
    </script>
    <%End If%>
</head>
<body>
    <script type="text/javascript" defer="defer">
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
                <img alt="Logo" src="../Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable">
            <%If TID = -1 Then%>
            <div id="BBSBMain">
                <%
            
                    CMD.CommandText = "SELECT * FROM [BBS_B] ORDER BY [ID]"
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    Do While DR.Read
                %>
                <div class="BBSB">
                    <a href="index.aspx?id=<%= DR("ID").ToString%>">
                        <img src="<%= DR("Logo").ToString %>" alt="<%= DR("BName").ToString %>" /></a>
                    <b><a href="index.aspx?id=<%= DR("ID").ToString%>">
                        <%= DR("BName").ToString %></a></b><br />
                    <%= DR("ReadMe").ToString%>
                    
                </div>
                <%
                
                Loop
                DR.Close()
                %>
            </div>
            <%Else%>
            <div class="BBSL">
                <div id="BBSPNew">
                    <input id="PNew" type="button" value="发新帖子" onclick="PNewP();" />
                </div>
                <div class="BBSB">
                    <img src="<%= GetB("Logo").ToString %>" alt="<%= GetB("BName").ToString %>" />
                    <b>
                        <%= GetB("BName").ToString%></b><br />
                    <%= GetB("ReadMe").ToString%>
                </div>
                <div class="Clear">
                </div>
                <div id="BBSList">
                    <div class="BBSLine">
                        <div class="BBSL1">
                            <b>标题</b></div>
                        <div class="BBSL2">
                            <b>发帖人</b></div>
                        <div class="BBSL3">
                            <b>最后回复时间</b></div>
                        <div class="BBSL4">
                            <b>回复/浏览</b></div>
                    </div>
                    <%
                        CMD.CommandText = "SELECT * FROM [BBS_P] WHERE [BID]=" & TID.ToString & " ORDER BY [LastTime] DESC"
                        Dim DR As OleDbDataReader = CMD.ExecuteReader
                        
                        Dim F = (ListPage - 1) * 50 + 1
                        Dim T = (ListPage - 1) * 50 + 1 + 50
                        Dim I = 1
                        Do While DR.Read And I < T
                            If I >= F And I < T Then
                    %>
                    <div class="BBSLine">
                        <div class="BBSL1">
                            <a href="p.aspx?id=<%= DR("ID").ToString%>">
                                <%= DR("Title").ToString%></a></div>
                        <div class="BBSL2">
                            <%= DR("UName").ToString%></div>
                        <div class="BBSL3">
                            <%= CString.MyTime(DR("LastTime").ToString)%></div>
                        <div class="BBSL4">
                            <%= New CBBSRe(DR("ID")).GetReNum()%>/<%= DR("PV").ToString%></div>
                    </div>
                    <%
                    End If
                    I = I + 1
                Loop
                DR.Close()
                    %>
                    <div class="Clear">
                    </div>
                </div>
                <%--翻页--%>
            </div>
            <%End If%>
        </div>
    </div>
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <br />
    <div id="Footer">
    </div>
    <script type="text/javascript" src="../JS/Frame.js"></script>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
</body>
</html>
