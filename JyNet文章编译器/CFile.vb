Imports Microsoft.VisualBasic
Imports System.IO
Public Class CFile
    Shared Function LoadFile(ByVal FilePath As String) As String
        Dim Line As String = ""
        Dim Sr As StreamReader = New StreamReader(FilePath, System.Text.Encoding.Default)
        Do While Sr.Peek() > 0
            Line = Line & Sr.ReadLine() & vbNewLine
        Loop
        LoadFile = Line
        Sr.Close()
        Sr = Nothing
    End Function
    Shared Sub WriteFile(ByVal FilePath As String, ByVal ConText As String)
        Dim SW As StreamWriter
        SW = File.CreateText(FilePath)
        SW.Write(ConText)
        SW.Close()
    End Sub
    Public Shared Function XMLKey(ByVal Source As String, ByVal Key As String) As String
        XMLKey = Mid(Source, InStr(Source, "<" & Key & ">") + Len(Key) + 2, InStr(Source, "</" & Key & ">") - InStr(Source, "<" & Key & ">") - Len(Key) - 2)
    End Function
    Public Shared Function XMLUp(ByVal Source As String, ByVal Key As String, ByVal Value As String) As String
        XMLUp = Left(Source, InStr(Source, "<" & Key & ">") + 1 + Len(Key)) & Value & Right(Source, Len(Source) - InStr(Source, "</" & Key & ">") + 1)
    End Function
  
End Class
