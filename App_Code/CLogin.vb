Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Imports System.Web.HttpContext
Public Class CLogin
    Public Shared Function GetUName() As String
        GetUName = ""
        If Not (Current.Request.Cookies("UserName") Is Nothing Or Current.Request.Cookies("PWD") Is Nothing) Then
            GetUName = Current.Request.Cookies("UserName").Value
        End If
    End Function
    Public Shared Sub LogOut()
        Current.Response.Cookies("UserName").Expires = DateTime.Now.AddDays(-1)
        Current.Response.Cookies("PWD").Expires = DateTime.Now.AddDays(-1)
    End Sub
    Public Shared Function Login(Optional ByVal UName As String = "null", Optional ByVal PWD As String = "null") As Integer
        If UName = "null" Then
            If Not (Current.Request.Cookies("UserName") Is Nothing Or Current.Request.Cookies("PWD") Is Nothing) Then
                UName = Current.Request.Cookies("UserName").Value
            End If
        End If
        If PWD = "null" Then
            If Not (Current.Request.Cookies("UserName") Is Nothing Or Current.Request.Cookies("PWD") Is Nothing) Then
                PWD = Current.Request.Cookies("PWD").Value
            End If
        End If
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand("select * from [Net_User] where [Username]='" & UName & "' and [PWD]='" & PWD & "'", CN)
        Dim DR As OleDbDataReader = CMD.ExecuteReader
        If DR.Read Then
            Login = 1
            '保存Cookie
            Current.Response.Cookies("UserName").Value = UName
            Current.Response.Cookies("UserName").Expires = DateTime.Now.AddDays(30)
            Current.Response.Cookies("PWD").Value = PWD
            Current.Response.Cookies("PWD").Expires = DateTime.Now.AddDays(30)
            Dim TU As CUser = New CUser(UName)
            '登录就送10金币
            If CType(TU.LastTime, DateTime).Day <> Now.Day Then
                '这里还有bug，如果是上个月的同一天将无法得到金币
                TU.Money = TU.Money + 10
            End If
            '刷新上次登录时间和IP


            TU.LastTime = Now.ToString
            TU.LastIP = Current.Request.UserHostAddress.ToString
        Else
            Login = 0
            If Not Current.Request.Cookies("PWD") Is Nothing Then
                Current.Request.Cookies("PWD").Expires = DateTime.Now.AddDays(-1)
            End If
        End If
        DR.Close()
        CMD.Dispose()
        CN.Close()

    End Function
End Class
