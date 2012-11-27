Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Imports System.Web.HttpContext
Public Class CMenu
    Shared Function Show() As String
        Show = ""
        Dim TM As Tag, tOut As String = ""
        For i = 1 To Tag.GetTagNum
            TM = New Tag(i)
            tOut = tOut & "<li class=" & Chr(34) & "Menu_others" & Chr(34) & "><a onmouseover=" & Chr(34) & "MShow(" & TM.ID & ");" & Chr(34) & " href=" & Chr(34) & "http://" & CString.DoMain & "/" & TM.PathName & Chr(34) & ">" & TM.Name & "</a></li>" & vbNewLine
        Next
        Current.Response.Write(Replace(CFile.LoadFile(Current.Server.MapPath("~/UI/MenuTem.htm")), "{jybox~main}", tOut))
    End Function
    Shared Function GetHtml() As String
        Dim TM As Tag, tOut As String = ""
        For i = 1 To Tag.GetTagNum
            TM = New Tag(i)
            tOut = tOut & "<li class=" & Chr(34) & "Menu_others" & Chr(34) & "><a onmouseover=" & Chr(34) & "MShow(" & TM.ID & ");" & Chr(34) & " href=" & Chr(34) & "http://" & CString.DoMain & "/" & TM.PathName & Chr(34) & ">" & TM.Name & "</a></li>" & vbNewLine
        Next
        GetHtml = Replace(CFile.LoadFile(Current.Server.MapPath("~/UI/MenuTem.htm")), "{jybox~main}", tOut)
    End Function
End Class
