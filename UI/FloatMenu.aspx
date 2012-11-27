<%@ Page Language="VB" %>

<%If Context.Request.QueryString("mode") = "post"Then%>
<%Select Case Int(Request.Form("id"))%>

<% Case 100%>
<a href="http://<%= CString.DoMain %>/index.aspx">主页</a> <a href="http://<%= CString.DoMain %>/Tag.aspx?tag=&num=30">最近更新</a> <a href="http://<%= CString.DoMain %>/Tag.aspx?tag=精华&num=30">精华文章</a> <a href="http://<%= CString.DoMain %>/jynet">本站源码</a> <a href="http://<%= CString.DoMain %>/Eula.aspx">用户协议&amp;声明</a> <a href="http://<%= CString.DoMain %>/BBS">论坛</a> <a href="http://<%= CString.DoMain %>/MyBlog">精英王子's Blog</a>
<%  Case 101%>
<a href="http://<%= CString.DoMain %>/soft/index.aspx">软件中心主页</a> <a href="http://<%= CString.DoMain %>/soft/open">开源软件推荐</a> <a href="http://<%= CString.DoMain %>/jynet">本站源码</a>
<%  Case 103%>
<a href="http://<%= CString.DoMain %>/BBS/index.aspx">论坛主页</a>
<%  Case 104%>
<a href="http://<%= CString.DoMain %>/My.aspx">个人中心</a> <a href="http://<%= CString.DoMain %>/Reg.aspx">注册帐号</a> <a href="http://<%= CString.DoMain %>/MyBlog">精英王子's Blog</a>
<%  Case 105%>
<a href="http://<%= CString.DoMain %>/MyBlog">精英王子's Blog</a>
<%  Case 1%>
<a href="http://<%= CString.DoMain %>/docs/index.aspx">技术教程主页</a> <a href="http://<%= CString.DoMain %>/docs/system">操作系统下载</a>
<%  Case 2%>
<a href="http://<%= CString.DoMain %>/program/index.aspx">编程技术主页</a> <a href="http://<%= CString.DoMain %>/program/tools">开发工具下载</a>
<%  Case 3%>
<a href="http://<%= CString.DoMain %>/read/index.aspx">文摘主页</a>


<%End Select%>
<%End If%>

