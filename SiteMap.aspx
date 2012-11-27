<?xml version="1.0" encoding="UTF-8" ?>
<%@ Page Language="VB" AutoEventWireup="false" CodeFile="SiteMap.aspx.vb" Inherits="SiteMap" %>
<% 
    Response.ContentType = "text/xml"
%>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
     <url>
       <loc>http://jybox.net/</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>always</changefreq>
       <priority>1.0</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/My.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.8</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/Reg.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.8</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/Tag.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.8</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/ad-disk.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/Eula.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.5</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/soft/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/soft/Open/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/JyNet/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/MyBlog/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.9</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/MyBlog/Art.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/MyBlog/About.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.8</priority>
     </url>

     <%--软件具体页--%>
     <url>
       <loc>http://<%= CString.DoMain %>/soft/Win7Full/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/soft/MiniTask/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>
     <url>
       <loc>http://<%= CString.DoMain %>/soft/MFC2010/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.6</priority>
     </url>

     <%--输出文章栏目主页--%>
     <%For i = 1 To Tag.GetTagNum%>
     <%TM = New Tag(i)%>
     <url>
       <loc>http://<%= CString.DoMain %>/<%= TM.PathName %>/index.aspx</loc>
       <lastmod><%= Now.AddHours(-25).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.9</priority>
     </url>
     <%Next%>

     <%--输出文章页--%>
     <% 
         CMD.CommandText = "select * from [Articles] ORDER BY [PV] DESC"
         Dim DR = CMD.ExecuteReader
         Do While DR.Read
     %>
     <url>
       <loc>http://<%= CString.DoMain %>/<%= New Tag(DR("Type")).PathName%>/<%=New CArt(DR("ID")).FileName %>.htm</loc>
       <lastmod><%= CType(New CArt(DR("ID")).LastTime, DateTime).ToString("s")%>+08:00</lastmod>
       <changefreq>daily</changefreq>
       <priority>0.5</priority>
     </url>
     <%  Loop
         DR.Close()
         
         %>
</urlset>
