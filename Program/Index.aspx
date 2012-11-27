<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Index.aspx.vb" Inherits="Program_Index" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>
        <%=New Tag(TypeID).Name%>_主页_精英盒子.Net</title>
        <meta name="keywords" content="<%= New Tag(TypeID).KeyWord.ToString%>,精英盒子,精英王子,Jybox" />
    <meta name="description" content="编程开发技术，VB、C++、VC、.Net" />
    <link href="../CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/WebDialog.css" rel="stylesheet" type="text/css" />

    <script src="../JS/User.js" type="text/javascript"></script>

    <script src="../JS/WebAjax.js" type="text/javascript"></script>

    <script src="../JS/ArtCL.js" type="text/javascript"></script>

    <script src="../JS/WebDialog.js" type="text/javascript"></script>

    <script src="../JS/WebCommon.js" type="text/javascript"></script>

    <link href="../CSS/ArtCL.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/Art10.css" rel="stylesheet" type="text/css" />
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
                <img alt="Logo" src="../Images/logo.png" />
            </div>
        </div>
        <%=CMenu.Show%>
        <div class="MainTable">
            <div id="TopTitle">
                <%=New Tag(TypeID).Name%>主页
            </div>
            <div id="ByKey">
                按关键字检索文章:
                <%
            
            
                    CMD.CommandText = "select * from [Art_Key] where [Type]=" & TypeID.ToString & " ORDER BY [PV] DESC"
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    I = 1
                    Do While DR.Read And I <= 10
                        Response.Write("<a target='_blank' href='http://" & CString.DoMain & "/Tag.aspx?tag=" & Server.UrlEncode(DR("Title").ToString) & "'>" & DR("Title") & "</a> &nbsp;")
                        I = I + 1
                    Loop
                    
            
                %>
                <input id="MyTag" type="text" /><input onclick="DoMyTag();" type="button" value="检索" />
            </div>
            <div id="ByMonth">
                <div class="Art10" style="width: 100%; height: auto;">
                    <h4>
                        月份归档</h4>
                    <ul>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.8'>2010.8</a></li>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.9'>2010.9</a></li>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.10'>2010.10</a></li>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.11'>2010.11</a></li>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2010.12'>2010.12</a></li>
                        <li><a target='_blank' href='http://jybox.net/Tag.aspx?tag=2011.1'>2011.1</a></li>
                    </ul>
                </div>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        CMD.CommandText = "select * from [Articles] where [Type]=" & TypeID.ToString & " ORDER BY [PV] DESC"
                        DR.Close()
                        DR = CMD.ExecuteReader
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
            
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
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
            
            
                        CMDN.CommandText = "select * from [Articles] where [Type]=" & TypeID.ToString & " ORDER BY [BuildTime] DESC"
                        Dim DRN As OleDbDataReader = CMDN.ExecuteReader
                        I = 1
                        Do While DRN.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DRN("ID")).FileName & ".htm'>" & New CArt(DRN("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                        
            
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                        
            
                    %>
                </ul>
            </div>
            <div class="Art10" style="float: left">
                <h4>
                    访问量最大文章</h4>
                <ul>
                    <%
            
            
                        I = 1
                        Do While DR.Read And I <= 10
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DR("ID")).FileName & ".htm'>" & New CArt(DR("ID")).Title & "</a></li>")
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
                            Response.Write("<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(TypeID).PathName & "/" & New CArt(DRN("ID")).FileName & ".htm'>" & New CArt(DRN("ID")).Title & "</a></li>")
                            I = I + 1
                        Loop
                        
                        DR.Close()
                        DRN.Close()
                    %>
                </ul>
            </div>
        </div>
    </div>
    <div class="Clear">
    </div>
    <div id="Footer">
    </div>
    </form>
    <!--#include file ="~/UI/ALLFooter.htm"-->
    <script type="text/javascript" src="../JS/Frame.js"></script>

</body>
</html>
