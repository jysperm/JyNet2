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
    ''' <summary>
    ''' ASP.NET提供文件下载函数（支持大文件、续传、速度限制、资源占用小）
    ''' </summary>
    ''' <param name="_Request">Page.Request对象</param>
    ''' <param name="_Response">Page.Response对象</param>
    ''' <param name="_fileName">下载文件名</param>
    ''' <param name="_fullPath">带文件名下载路径</param>
    ''' <param name="_speed">每秒允许下载的字节数</param>
    ''' <returns>返回是否成功</returns>
    ''' <remarks>
    ''' '输出硬盘文件，提供下载 调用例
    '''</remarks>
    Public Shared Function ResponseFile(ByVal _Request As HttpRequest, ByVal _Response As HttpResponse, ByVal _fileName As String, ByVal _fullPath As String, ByVal _speed As Long) As Boolean
        Try
            Dim myFile As New System.IO.FileStream(_fullPath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
            Dim br As New System.IO.BinaryReader(myFile)
            Try
                _Response.AddHeader("Accept-Ranges", "bytes")
                _Response.Buffer = False
                Dim fileLength As Long = myFile.Length
                Dim startBytes As Long = 0

                Dim pack As Integer = 10240
                '10K bytes 
                'int sleep = 200; //每秒5次 即5*10K bytes每秒 
                Dim sleep As Integer = CInt(Math.Floor(1000 * pack / _speed)) + 1
                If _Request.Headers("Range") IsNot Nothing Then
                    _Response.StatusCode = 206
                    Dim range As String() = _Request.Headers("Range").Split(New Char() {"="c, "-"c})
                    startBytes = Convert.ToInt64(range(1))
                End If
                _Response.AddHeader("Content-Length", (fileLength - startBytes).ToString())
                If startBytes <> 0 Then
                    _Response.AddHeader("Content-Range", String.Format(" bytes {0}-{1}/{2}", startBytes, fileLength - 1, fileLength))
                End If
                _Response.AddHeader("Connection", "Keep-Alive")
                _Response.ContentType = "application/octet-stream"
                _Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(_fileName, System.Text.Encoding.UTF8))

                br.BaseStream.Seek(startBytes, System.IO.SeekOrigin.Begin)
                Dim maxCount As Integer = CInt(Math.Floor((fileLength - startBytes) / pack)) + 1

                For i As Integer = 0 To maxCount - 1
                    If _Response.IsClientConnected Then
                        _Response.BinaryWrite(br.ReadBytes(pack))
                        System.Threading.Thread.Sleep(sleep)
                    Else
                        i = maxCount
                    End If
                Next
            Catch
                Return False
            Finally
                br.Close()
                myFile.Close()
            End Try
        Catch
            Return False
        End Try
        Return True
    End Function
End Class
