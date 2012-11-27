Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Public Class CMenu
    Shared Function GetHtml() As String
        Dim TM As Tag, tOut As String = ""
        For i = 1 To Tag.GetTagNum
            TM = New Tag(i)
            tOut = tOut & "<li class=" & Chr(34) & "Menu_others" & Chr(34) & "><a onmouseover=" & Chr(34) & "MShow(" & TM.ID & ");" & Chr(34) & " href=" & Chr(34) & "http://" & CString.DoMain & "/" & TM.PathName & Chr(34) & ">" & TM.Name & "</a></li>" & vbNewLine
        Next
        GetHtml = Replace(CFile.LoadFile("../UI/MenuTem.htm"), "{jybox~main}", tOut)
    End Function
End Class
