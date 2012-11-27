Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Imports System.Web.HttpContext
Public Class CUser
    Public P_UName As String
    Sub New(ByVal UName As String)
        P_UName = UName
    End Sub
    Property UName() As String
        Get
            UName = P_UName
        End Get
        Set(ByVal value As String)
            P_UName = value
        End Set
    End Property
    ReadOnly Property UID() As Integer
        Get
            UID = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                UID = DR("ID").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
    End Property
    Property PWD() As String
        Get
            PWD = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                PWD = DR("PWD").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [PWD] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property NName() As String
        Get
            NName = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                NName = DR("NName").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [NName] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property RegTime() As String
        Get
            RegTime = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                RegTime = DR("RegTime").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [RegTime] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property RegIP() As String
        Get
            RegIP = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                RegIP = DR("RegIP").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [RegIP] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property QQ() As String
        Get
            QQ = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                QQ = DR("QQ").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [QQ] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Exp() As Integer
        Get
            Exp = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Exp = DR("Exp").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [Exp] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Lv() As Integer
        Get
            Lv = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Lv = DR("Lv").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [Lv] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Money() As Integer
        Get
            Money = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Money = DR("Money").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [Money] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Logo() As String
        Get
            Logo = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Logo = DR("Logo").ToString
            End If
            If Left(Logo, 4).ToUpper <> "HTTP" Then
                Logo = "http://jybox.net" & Logo
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [Logo] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property EMail() As String
        Get
            EMail = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                EMail = DR("EMail").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [EMail] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property LastIP() As String
        Get
            LastIP = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                LastIP = DR("LastIP").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [LastIP] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property LastTime() As String
        Get
            LastTime = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_User] where [UserName]='" & P_UName.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                LastTime = DR("LastTime").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Net_User] SET [LastTime] = '" & value.ToString & "' WHERE [UserName] = '" & P_UName.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Shared Function Reg(ByVal UName As String, ByVal PWD As String, ByVal NName As String, ByVal Mail As String) As String
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select * from [Net_User] where [UserName]='" & UName.ToString & "'"
        Dim DR As OleDbDataReader = CMD.ExecuteReader
        If DR.Read Then                                                            '当能从DataReader中读出数据时，说明用户名重复
            Reg = "sameuser"
        Else
            DR.Close()
            CMD.CommandText = "INSERT INTO [Net_User] (UserName, NName, PWD,Email,RegIP,LastIP) VALUES(@UserName, @NName, @PWD,@Email,@RegIP,@LastIP)"
            '向SQL语句中补充数据
            CMD.Parameters.AddWithValue("@UserName", UName.ToString)
            CMD.Parameters.AddWithValue("@NName", NName.ToString)
            CMD.Parameters.AddWithValue("@PWD", CString.MD5(PWD.ToString))         '对密码MD5后存入数据库
            CMD.Parameters.AddWithValue("@Email", Mail.ToString)
            CMD.Parameters.AddWithValue("@RegIP", Current.Request.UserHostAddress.ToString)
            CMD.Parameters.AddWithValue("@LastIP", Current.Request.UserHostAddress.ToString)
            '执行命令
            CMD.ExecuteNonQuery()
            '重新复核一下是否注册成功
            CMD.CommandText = "select * from [Net_User] where [Username]='" & UName.ToString & "'"
            DR = CMD.ExecuteReader()
            If DR.Read Then
                Reg = "OK"
                Current.Response.Cookies("UserName").Value = UName
                Current.Response.Cookies("UserName").Expires = DateTime.Now.AddDays(30)
                Current.Response.Cookies("PWD").Value = PWD
                Current.Response.Cookies("PWD").Expires = DateTime.Now.AddDays(30)
            Else
                Reg = "ENF"
            End If
        End If
        DR.Close()
        CMD.Dispose()
        CN.Close()
    End Function
End Class
