<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script runat="server">

</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>声明&协议_精英盒子.Net</title>
    <meta name="keywords" content="用户,声明,协议,JyNet,精英王子,精英盒子,精英盒子.Net" />
    <link href="CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="CSS/Reg.css" rel="stylesheet" type="text/css" />
    <link href="CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="JS/WebAjax.js" type="text/javascript"></script>
    <script src="JS/WebDialog.js" type="text/javascript"></script>
    <script src="JS/WebCommon.js" type="text/javascript"></script>
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
        <div id="HeadBar">
            加载中...
        </div>
        <div class="HeadArea">
            <div class="HeadLogo">
                <img alt="Logo" src="Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable" style="text-align: left">
            <h1>
                声明&协议</h1>
                <span style="color: #FF0000">注：本声明是针对精英盒子.Net的，本站的开源版本命名为JyNet，拥有另一套协议</span>
            <ol>
                <li>精英盒子.Net主体(不包括论坛)中的文章，如未特殊注明，即为精英盒子.Net原创<br />
                    欢迎转载，转载请保留原链接。您可以对文章进行修改，但不能违反文章的本意，修改请注明来自精英盒子.Net<br />
                    如果您对文章讨论的内容有疑问，可以留言询问，或发邮件至m@jybox.net</li>
                <li>对于明确注明作者的文章，有任何疑问请联系作者<br />
                    如果您是作者，不希望您的文章被转载到精英盒子.Net，请联系我们(m@jybox.net)</li>
                <li>精英盒子.Net提供的所有小工具均为精英盒子.Net原创，任何问题请联系我们(m@jybox.net) </li>
                <li>精英盒子.Net所附的下载（资料/软件），即使是我们制作的，也和精英盒子.Net没有直接关系<br />
                    其他（非我们制作的）软件/资料，如果您是作者，不希望我们提供下载，请联系我们(m@jybox.net)<br />
                    对于精英盒子.Net所附的下载（资料/软件）的任何问题，与精英盒子.Net无关，请单独联系其作者</li>
                <li>对于精英盒子.Net的论坛、文章评论功能，评论的内容（即使是“我”评论的），与精英盒子.Net无关，评论内观点不直接代表精英盒子.Net的立场<br />
                    精英盒子.Net将先允许用户进行（发表/评论）操作，然后进行审核，去除不合适的言论，情节严重的将作处罚</li>
                <li>精英盒子.Net所使用的程序(包括网站主体、论坛、小工具、作者博客)全部为精英盒子.Net自行开发，没有使用其他任何同类程序的文件或代码片段。因此，精英盒子.Net所使用的一切程序版权属于精英盒子.Net的开发者们。<br />
                    任何人未经允许不得引用精英盒子.Net的任何代码（包括输出的HTML、CSS样式、JS文件/JS代码段、图片、服务器代码等）<br />
                    未经允许不得引用精英盒子.Net的图片或文件的下载链接(俗称“盗链”)<br />
                    如希望获得允许或希望购买网站源文件或购买网站，请联系我们（m@jybox.net）</li>
                <li>精英盒子.Net的用户不得以任何理由利用精英盒子.Net存在的bug<br />
                    发现bug最好及时与我们联系（m@jybox.net）
                    <br />
                    用户不得参与任何破坏/影响精英盒子.Net网站运行的活动<br />
                    包括但不限于：<br />
                    大量或高速，或大量And高速刷新页面或点击链接，尝试修改Get/Post/Cookies参数,尝试访问未明确公开的目录或文件,尝试打开或下载源码文件和数据库文件,尝试输入明显错误的数据来故意导致错误的发生<br />
                    严重者将作处罚</li>
                <li>在内测期间(版本号小于2.10)，我们将随时删除、回滚、修改存档，不会做任何通知和补偿</li>
            </ol>
            
        </div>
    </div>
    <div id="Footer">
    </div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="JS/Frame.js"></script>
</body>
</html>
