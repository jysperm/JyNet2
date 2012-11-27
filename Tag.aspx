<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Tag.aspx.vb" Inherits="_Tag" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>文章检索_精英盒子.Net</title>
    <link href="CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="JS/User.js" type="text/javascript"></script>
    <script src="JS/WebAjax.js" type="text/javascript"></script>
    <link href="CSS/Tag.css" rel="stylesheet" type="text/css" />
    <script src="JS/WebDialog.js" type="text/javascript"></script>
    <script src="JS/WebCommon.js" type="text/javascript"></script>
    <link href="CSS/Art10.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <%
        sTag = IIf(Request.QueryString("tag") = Nothing, "", Request.QueryString("tag"))
        sOrder = IIf(Request.QueryString("order") = Nothing, "time", Request.QueryString("order"))
        sNum = IIf(Request.QueryString("num") = Nothing, "100", Request.QueryString("num"))
    
    %>
    <script type="text/javascript">
        var dialog = new WebDialog();
        dialog.ImgZIndex = 107;
        dialog.DialogZIndex = 108;
        dialog.Width = 400;
        var Order = "<%=sOrder %>"
    </script>
    <form id="FormMain">
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
            <div id="ByKey">
                按关键字检索文章:
                <%
            
            
                    CMD.CommandText = "select * from [Art_Key] ORDER BY [PV] DESC"
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    I = 1
                    Do While DR.Read And I <= 20
                        Response.Write("<a href='http://" & CString.DoMain & "/Tag.aspx?tag=" & Server.UrlEncode(DR("Title").ToString) & "'>" & DR("Title") & "</a> &nbsp;")
                        I = I + 1
                    Loop
                    DR.Close()
            
                %>
                <input id="MyTag" type="text" /><input onclick="DoTag();" type="button" value="检索" />
            </div>
            <div id="Set">
                <b>选项：</b><br />
                排序依据<input type="button" value="时间(从新到旧)" onclick="Order='time';DoTag();" /><input type="button" value="浏览量(从大到小)" onclick="Order='pv';DoTag();" /><br />
                显示数量:<input id="Num" type="text" /><br />
                <input type="button" value="确定" onclick="DoTag();" />
            </div>
            <div id="Out">
            <b><span style="color: #0066FF">搜索结果</span> ---- 您搜索的关键词是<%= sTag%>，一次显示<%= sNum%>篇文章，按<%= IIf(sOrder = "pv", "时间(从新到旧)", "浏览量(从大到小)")%>排序</b>
            <ul>
                <%
            
                    
                    CMD.CommandText = "select * from [Articles] WHERE [KeyWord] LIKE '%" & sTag & "%' ORDER BY [" & IIf(sOrder = "pv", "PV", "BuildTime") & "] DESC"
                    DR = CMD.ExecuteReader
                    I = 1
                    Do While DR.Read And I <= Int(sNum)
                        Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DR("Type")).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
                        I = I + 1
                    Loop
                    DR.Close()
            
                %>
                <%If I = 1 Then%>
                悲剧啊，竟然一篇相关的文章都没有
                <% End If%>
                </ul>
            </div>
            <div id="D">
            <b>一些提示：</b><br />
            1.请尽量使用简短的词语来进行搜索，本搜索是完全匹配的，不支持模糊查询<br />
            2.请不要设置一次显示太多文章(超过200)，否则可能导致浏览器卡死<br />
            3.同时欢迎大家向我们投稿，请发信到m@jybox.net<br />
            <a href='http://jybox.net/SiteMap.aspx'>给机器人看的网站地图</a>
            </div>
        </div>
        <div class="Clear">
        </div>
        <div id="Footer">
        </div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="JS/Frame.js"></script>
    <script type="text/javascript">
        function TLoad() {
            $("Num").value = "<%=sNum %>";
            $("MyTag").value = "<%=sTag %>";
        }
        function DoTag() {
            window.location = "Tag.aspx?tag=" + encodeURIComponent($("MyTag").value) + "&num=" + $("Num").value + "&order=" + Order;
        }
        TLoad();
    </script>
</body>
</html>
