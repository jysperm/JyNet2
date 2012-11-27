Imports Microsoft.VisualBasic
Public Class CString
    Public Shared DoMain As String = "jybox.net"
    Public Shared Function SameString(ByVal Source As String, ByVal Num As Integer) As String
        SameString = ""
        For i = 1 To Num
            SameString &= Source
        Next
    End Function
    Public Shared Function MD5(ByVal strSource As String, Optional ByVal Code As Int16 = 16) As String
        '这里用的是ascii编码密码原文，如果要用汉字做密码，可以用UnicodeEncoding，但会与ASP中的MD5函数不兼容
        Dim dataToHash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(strSource)
        Dim hashvalue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(dataToHash)
        Dim i As Integer
        MD5 = ""
        Select Case Code
            Case 16  '选择16位字符的加密结果
                For i = 4 To 11
                    MD5 += Hex(hashvalue(i)).ToLower
                Next
            Case 32  '选择32位字符的加密结果
                For i = 0 To 15
                    MD5 += Hex(hashvalue(i)).ToLower
                Next
            Case Else   'Code错误时，返回全部字符串，即32位字符
                For i = 0 To hashvalue.Length - 1
                    MD5 += Hex(hashvalue(i)).ToLower
                Next
        End Select
    End Function
    Public Shared Function MyTime(ByVal TheTime As Date) As String
        Dim tSNum = DateDiff("s", TheTime, Now)
        Select Case tSNum
            Case 0 To 59
                MyTime = tSNum.ToString & "秒前"
            Case 60 To 1799
                MyTime = Int(tSNum / 60).ToString & "分钟前"
            Case 1800 To 1859
                MyTime = "半小时前"
            Case 1860 To 3599
                MyTime = Int(tSNum / 60).ToString & "分钟前"
            Case 3600 To 3659
                MyTime = "一小时前"
            Case 3660 To 86399
                MyTime = Int(tSNum / 3600).ToString & "小时前"
            Case 86400 To 90000
                MyTime = "一天前"
            Case 90000 To 172800
                MyTime = "昨天 " & TheTime.GetDateTimeFormats(CChar("t"))(0).ToString()
            Case 172801 To 259200
                MyTime = "前天 " & TheTime.GetDateTimeFormats(CChar("t"))(0).ToString()
            Case Else
                If tSNum < 31557600 Then
                    MyTime = Format(TheTime, "M-d")
                Else
                    MyTime = Format(TheTime, "yyyy-M-d")
                End If
        End Select
    End Function
    Public Shared Function EnSym(ByVal strDataIn As String, ByVal Key As String) As String
        Dim intXORValue1 As Integer
        Dim intXORValue2 As Integer
        Dim strDataOut As String = ""
        Mid(Key, 2, 1) = Asc(Chr(1))
        For i = 1 To Len(strDataIn)
            intXORValue1 = Asc(Mid$(strDataIn, i, 1))
            intXORValue2 = Asc(Mid$(Key, ((i Mod Len(Key)) + 1), 1))
            strDataOut = strDataOut + Chr(intXORValue1 Xor intXORValue2)
        Next i
        EnSym = strDataOut
    End Function
End Class
