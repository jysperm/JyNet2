<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Art.aspx.vb" Inherits="MyBlog_Art" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <%If ArtList = 1 Then%>
    <title>
        <%= GetArt("Title").ToString%>_精英王子's Blog</title>
    <%Else%>
    <title>文章列表_精英王子's Blog</title>
    <%End If%>
    <%If ArtList = 1 Then%>
    <meta name="keywords" content="<%= GetArt("KeyWords").ToString%>,精英王子,精英盒子,王子亭,Jybox,Jingyingbox,精英,JY,王子,个人,博客,Blog" />
    <%Else%>
    <meta name="keywords" content="精英王子,精英盒子,王子亭,Jybox,Jingyingbox,精英,JY,王子,个人,博客,Blog" />
    <%End If%>
    <meta name="description" content="这里是精英王子的博客，用来记录一些点滴，不写技术性文章" />
    <link href="../CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebAjax.js" type="text/javascript"></script>
    <link href="../CSS/MyBlog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebDialog.js" type="text/javascript"></script>
    <script src="../JS/WebCommon.js" type="text/javascript"></script>
    <script src="../JS/User.js" type="text/javascript"></script>
    
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
        <div id="HeadArea">
            <div id="BMenu">
                <ul style="margin: 0px 0px 0px 0px">
                    <li><a href="http://jybox.net">网站主页</a></li>
                    <li><a href="http://jybox.net/MyBlog">博客主页</a></li>
                    <li><a href="http://jybox.net/MyBlog/Art.aspx">日志列表</a></li>
                    <li><a href="http://jybox.net/MyBlog/About.aspx">关于我</a></li>
                </ul>
            </div>
        </div>
        <div class="MainTable">
            <%If ArtList = 1 Then%>
            <div id="ArtMain">
                <div class="ATitle">
                    <%= GetArt("Title").ToString%>
                </div>
                <div class="AInfo">
                    <%= GetArt("TheTime").ToString%>(<%= CString.MyTime(GetArt("TheTime").ToString)%>)
                    | 访问量:
                    <%= GetArt("PV").ToString%>
                    PageView
                </div>
                <div class="ArtBody">
                    <%= GetArt("Body").ToString%>
                </div>
                <div class="ArtFooter">
                    访问量 :
                    <%= GetArt("PV").ToString%>
                    PageView
                    <br />
                    上一篇:<a href="http://<%=CString.DoMain%>/MyBlog/Art.aspx?id=<%= GetArt("ID",AID - 1).ToString%>"><%= GetArt("Title", AID - 1).ToString%></a>
                    |下一篇:
                    <%If AID < GetArts() Then%>
                    <a href="http://<%=CString.DoMain%>/MyBlog/Art.aspx?id=<%=GetArt("ID",AID + 1).ToString%>">
                        <%= GetArt("Title", AID + 1).ToString%></a>
                    <%Else%>
                    没有了！
                    <%End If%>
                </div>
                <div id="Say">
                    评论加载中...
                </div>
                <div id="SayNew">
                    <a id="ShowAll" onclick="Show(100);return false;" href="#">以上是最后一部分评论，点击这里查看全部</a>
                    <br />
                    <hr />
                    <%For I = 1 To 15%>
                    <img onclick="AddEm(<%= I %>);" src="../Images/EM/<%= I %>.gif" class="Em" />
                    <%Next%>
                    <span id="Info">正在加载用户信息...</span>
                    <textarea id="In"></textarea>
                    <div id="VCodeIn">
                        验证码：<a href="#" onclick="ReVCode();return false;"><img id="VCodeImg" alt="点击更换验证码"
                            src="../UI/VCode.aspx" /></a>
                        <input id="VCode" type="text" /><a href="#" onclick="Login(ALoad);" style="font-size: 13px">登录后无需验证码</a>
                    </div>
                    <input id="OK" type="button" value="发表" onclick="NewSay();" />
                </div>
            </div>
            <%Else%>
            <div id="ArtList">
                <h2>
                    文章列表</h2>
                <%
                    Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                    Dim CMD As New OleDbCommand
                    CMD.Connection = CN
                    CMD.CommandText = "select * from [Blog_Art] ORDER BY [ID] DESC"
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    Dim F = (ListPage - 1) * 20 + 1
                    Dim T = (ListPage - 1) * 20 + 1 + 20
                    Dim I = 1
                    Do While DR.Read And I < T
                        If I >= F And I < T Then
                %>
                <div class="ATitle">
                    <a href="http://<%=CString.DoMain%>/MyBlog/Art.aspx?id=<%=DR("ID").ToString%>">
                        <%= DR("Title").ToString%>
                    </a>
                </div>
                <div class="AInfo">
                    <%= DR("TheTime").ToString%>(<%= CString.MyTime(DR("TheTime").ToString)%>) | 访问量:
                    <%= DR("PV").ToString%>
                    PageView
                </div>
                <%
                End If
                I = I + 1
            Loop
            DR.Close()
            CMD.Dispose()
            CN.Close()
                
                %>
            </div>
            <%End If%>
        </div>
    </div>
    </form>
    <div id="F">
        <br />
        <br />
        <br />
        By 精英王子 2009-2011
    </div>
    <script type="text/javascript" defer="defer">
    <%If ArtList = 1 Then%>
    var id=<%= GetArt("ID").ToString%>;
    <%End If%>
        
        var dialog = new WebDialog();
        document.body.oncopy=function(){setTimeout(function(){var text=clipboardData.getData("text");if(text){text=text+"\r\n本文来自: 精英盒子.Net(jybox.net) ，欢迎来访！详细出处请参考："+location.href;clipboardData.setData("text",text)}},100)}
        function left(mainStr,lngLen) { if (lngLen>0) {return mainStr.substring(0,lngLen)} else{return null} } 
        function right(mainStr,lngLen) { if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) { return mainStr.substring(mainStr.length-lngLen,mainStr.length)} else{return null} } 
        function mid(mainStr,starnum,endnum){ if (mainStr.length>=0){ return mainStr.substr(starnum,endnum) }else{return null} } 

        function Root(u)
        {
            var p=location.pathname;
            for(var r="",i=1,t=0;i<p.length;i++,t++)
                if(mid(p,i,1)=="/" && t>1)
                    r=r+"../";
            return r+u
        }

        function insertAtCursor(myField, myValue) {
            if (document.selection) {
                myField.focus();
                sel = document.selection.createRange();
                sel.text = myValue;
                sel.select();
            }
            else if (myField.selectionStart || myField.selectionStart == '0') {
                var startPos = myField.selectionStart;
                var endPos = myField.selectionEnd;
                var restoreTop = myField.scrollTop;
                myField.value = myField.value.substring(0, startPos) + myValue + myField.value.substring(endPos, myField.value.length);
                if (restoreTop > 0) {
                    myField.scrollTop = restoreTop;
                }
                myField.focus();
                myField.selectionStart = startPos + myValue.length;
                myField.selectionEnd = startPos + myValue.length;
            } else {
                myField.value += myValue;
                myField.focus();
            }
        }
        function AddEm(n) {
            insertAtCursor($("In"), "[/表情" + n + "]");
        }
        function ReVCode() {
            times++;
            document.getElementById("VCodeImg").src = "";
            document.getElementById("VCodeImg").src = "../UI/VCode.aspx?times=" + times;
        }

        function ALoad() {
            var result = executeHttpRequest("POST", "../AJAX/Blog_Say.ashx?mode=post", "do=getlogin");
            Show(10);
            if (result == "null") {
                $("Info").innerHTML = "<span style='color:red;'>[您无需登录也可以发表评论!]</span>&nbsp;&nbsp;<a href='#' onclick='Login(ALoad);' style='font-size: 13px'>[点击这里登录(无需刷新)]</a>";
            }
            else {
                $("Info").innerHTML = "[" + result + "]写下你的看法吧！";
                $("VCodeIn").style.display = "none";
            }
        }
        var times = 1;
        ALoad();
        function NewSay() {
            if ($("In").value == "") {
                alert("请输入评论内容");
                return;
            }
            result = executeHttpRequest("POST", "../AJAX/Blog_Say.ashx?mode=post", "do=new&aid=" + id + "&body=" + encodeURIComponent($("In").value) + "&vcode=" + $("VCode").value);
            ReVCode();
            if (result == "OK") {
                alert("评论成功");
                $("In").value = "";
                Show(10);
            }
            else {
                alert(result);
            }
        }
        function Show(n) {
            $("Say").innerHTML = executeHttpRequest("POST", "../AJAX/Blog_Say.ashx?mode=post", "do=get&num=" + n + "&aid=" + id);
        }
    </script>
    <script type="text/javascript">
        var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
        document.write(unescape("%3Cscript src='" + _bdhmProtocol + "hm.baidu.com/h.js%3F77016691cd5a049005dba568b5164b59' type='text/javascript'%3E%3C/script%3E"));
    </script>
</body>
</html>
