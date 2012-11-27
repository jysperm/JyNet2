<%@ Page Language="VB" CodeFile="Index.aspx.vb" Inherits="Index" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <meta name="google-site-verification" content="HkqekEnUPCdKdS6J7JjAiu5wICdzjwhgysUIU8JQcXo" />
    <title>精英盒子.Net_主页_JyNet_精英王子的开源软件工作室</title>
    <meta name="keywords" content="编程,文摘,主页,JyNet,精英王子,精英盒子,精英盒子.Net,开源" />
    <meta name="description" content="精英盒子.Net是由精英王子自主研发，面向编程开发/计算机爱好者的综合性网站，同时发布一些自主研发的开源程序" />
    <link href="CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="JS/User.js" type="text/javascript"></script>
    <script src="JS/WebAjax.js" type="text/javascript"></script>
    <link href="CSS/Index.css" rel="stylesheet" type="text/css" />
    <script src="JS/WebDialog.js" type="text/javascript"></script>
    <script src="JS/WebCommon.js" type="text/javascript"></script>
    <link href="CSS/Art10.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <script type="text/javascript">
        function DoMyTag() {
            window.location = "Tag.aspx?tag=" + encodeURIComponent($("MyTag").value);
        }
        function DoLogOut() {
            var result = executeHttpRequest("POST", "AJAX/Login.ashx?mode=post", "do=logout");
            window.location = "Index.aspx";
        }
        function ILoad() {
            window.location = "Index.aspx";
        }
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
            <div id="DivTL">
                <div id="ByKey">
                    按关键字检索文章:
                    <%
            
            
                        CMD.CommandText = "select * from [Art_Key] ORDER BY [PV] DESC"
                        Dim DR As OleDbDataReader = CMD.ExecuteReader
                        I = 1
                        Do While DR.Read And I <= 20
                            Response.Write("<a target='_blank' href='http://" & CString.DoMain & "/Tag.aspx?tag=" & Server.UrlEncode(DR("Title").ToString) & "'>" & DR("Title") & "</a> &nbsp;")
                            I = I + 1
                        Loop
                    
            
                    %>
                    <input id="MyTag" type="text" /><input onclick="DoMyTag();" type="button" value="检索" />
                </div>
                <div id="Focus">
                    <!--焦点↓-->
                    <div style="margin: 20px 10px 20px 10px; font-size: 17px; font-weight: bold;">
                        精英盒子论坛开发完成，<a href="bbs">>>点击这里进入<<</a>！<br />
                        
                    </div>
                    <!--焦点↑-->
                </div>
            </div>
            <div id="DivTR">
                <div id="DLogin">
                    <% If IsLoged Then%>
                    <div id="UIn">
                        <img id="Logo" alt="用户头像" src="<%=TU.Logo.ToString%>" />
                        <span class="NName" id="NName">
                            <%=TU.NName.ToString%></span><br />
                        <img class="ImgMoney" src="Images/Money.jpg" /><span class="Money"><%=TU.Money.ToString%></span><br />
                        <span class="Exp">积分：<span class="ExpIn"><%=TU.Exp.ToString%></span></span>
                        <div class="Clear">
                        </div>
                        <input id="Pay" type="button" value="充值" />
                        <input type="button" value="个人中心" onclick="window.location = 'My.aspx'" style="float: left" />
                        <input id="LogOut" type="button" value="退出" onclick="DoLogOut();" />
                        <div class="Clear">
                        </div>
                    </div>
                    <%Else%>
                    <div style="margin: 30px 50px 0px 50px">
                        <h3>
                            登录精英盒子吧！</h3>
                        <input type="button" value="登录" onclick="Login(ILoad);" /><input type="button" value="注册"
                            onclick="window.location = 'Reg.aspx'" />
                    </div>
                </div>
                <%End If%>
                <div id="ABout">
                    联系我们：<br />
                    邮箱：m@jybox.net(<a href="mailto:m@jybox.net">发送邮件</a>)<br />
                    QQ：172157947<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=172157947&site=qq&menu=yes"><img
                        src="http://wpa.qq.com/pa?p=2:172157947:41" alt="点击这里给我发消息" style="vertical-align: middle;"></a><br />
                    QQ群：61137226<a target="_blank" href="http://qun.qq.com/#jointhegroup/gid/61137226"><img
                        src="Images/QQun.jpg" alt="点击这里加入此群" style="vertical-align: middle;"></a><br />
                </div>
            </div>
        </div>
        <div class="Clear">
        </div>
        <div id="DivDL">
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        CMD.CommandText = "select * from [Articles] ORDER BY [PV] DESC"
                        DR.Close()
                        DR = CMD.ExecuteReader
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DR("Type")).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
            
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    最新文章</h4>
                <ul>
                    <%
                        CMDN.CommandText = "select * from [Articles] ORDER BY [BuildTime] DESC"
                        Dim DRN As OleDbDataReader = CMDN.ExecuteReader
                        I = 1
                        Do While DRN.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DRN("Type")).PathName & "/" & New CArt(DRN("ID")).FileName & ".htm'>" & New CArt(DRN("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    论坛热帖</h4>
                <ul>
                <% 
                        CMDB.CommandText = "select * from [BBS_P] ORDER BY [LastTime] DESC"
                        Dim DRB As OleDbDataReader = CMDB.ExecuteReader
                        I = 1
                        Do While DRB.Read And I <= 10
                        Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/BBS/p.aspx?id=" & DRB("ID").ToString & "'>" & DRB("Title").ToString & "</a></li>")
                            I = I + 1
                    Loop
                    DRB.Close()
                %>
                </ul>
            </div>
            <div class="Clear">
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DR("Type")).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
            
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    最新文章</h4>
                <ul>
                    <%
                        I = 1
                        Do While DRN.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DRN("Type")).PathName & "/" & New CArt(DRN("ID")).FileName & ".htm'>" & New CArt(DRN("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    最新文章</h4>
                <ul>
                    <%
                        I = 1
                        Do While DRN.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(DRN("Type")).PathName & "/" & New CArt(DRN("ID")).FileName & ".htm'>" & New CArt(DRN("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                        DR.Close()
                        DRN.Close()
                    %>
                </ul>
            </div>
        </div>
        <div id="DivDR">
            <div class="Art10" style="width: 100%; height: auto;">
                <h4>
                    月份归档</h4>
                <ul>
                <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2011.2'>2011.2</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2011.1'>2011.1</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.12'>2010.12</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.11'>2010.11</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.10'>2010.10</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.9'>2010.9</a></li>
<li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.8'>2010.8</a></li>

                    
                    
                    
                    
                    
                    
                </ul>
            </div>
            <div class="Art10" style="width: 100%; height: auto;">
                <h4>
                    精英盒子.Net</h4>
                <ul>
                    <li><a>
                        <%= CFile.XMLKey(CFile.LoadFile(Server.MapPath("~/Admin/Setting.ini")), "Ver")%></a></li>
                    <li><a>
                        <%= CFile.XMLKey(CFile.LoadFile(Server.MapPath("~/Admin/Setting.ini")), "Last")%></li>
                    <li><a>By 精英王子</a></li>
                    <li><a>m@jybox.net</a></li>
                    <li><a href='http://jybox.net/SiteMap.aspx'>给机器人看的网站地图</a></li>
                    <li><a href='http://jybox.net/ad-disk.aspx'>广告联盟-网盘推荐</a></li>
                    <li><a href='http://jybox.net/Eula.aspx'>用户协议&声明</a></li>
                </ul>
            </div>
        </div>
        <div class="Clear">
        </div>
        <div id="Link">
            <b>友情链接</b> <a href="http://www.baidu.com">百度</a> <a href="http://www.uudisc.com/user/jingyingbox">
                UuShare网盘</a> <a href="http://hi.baidu.com/f4tb0y_">f4tb0y's blog</a> <a href="http://blog.sina.com.cn/u/1806240760">
                    limit's Blog</a> <a href="http://www.1314dhw.com/">导航屋编程论坛</a> <a href="#" onclick="alert('要做友链的速度，精英盒子.Net新版全新上线\n全心自行研发的系统、原创的内容，绝对有潜力。\n\n友链要求：在百度、谷歌都有收录，网站有实际内容，网络延迟1500以下，符合国家各项法律。\n同时，“软件中心”接受原创软件。\n软件要求：必须原创软件，免费软件，无侵权，符合国际各项法律。开源最好\n\n详询m@jybox.net\n')">
                            友链申请要求</a>
        </div>
    </div>
    <div class="Clear">
    </div>
    <div style="text-align: center">
        <%--页脚--%>
        <%= Replace(CFile.LoadFile(Server.MapPath("~/UI/Footer.htm")), "[Ver]", CFile.XMLKey(CFile.LoadFile(Server.MapPath("~/Admin/Setting.ini")), "Ver"))%>
    </div>
    <div id="Footer" style="visibility: hidden">
    </div>
    </div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="JS/Frame.js"></script>
</body>
</html>
