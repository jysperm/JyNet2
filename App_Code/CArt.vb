Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Imports System.Web.HttpContext
Public Class CArt
    Public P_ID As String
    Sub New(ByVal ID As String)
        P_ID = ID
    End Sub
    Property ID() As Integer
        Get
            ID = P_ID
        End Get
        Set(ByVal value As Integer)
            P_ID = value
        End Set
    End Property
    Property PV() As Integer
        Get
            PV = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                PV = DR("PV").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [PV] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Type() As Integer
        Get
            Type = 0
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Type = DR("Type").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Type] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property KeyWord() As String
        Get
            KeyWord = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                KeyWord = DR("KeyWord").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [KeyWord] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Describe() As String
        Get
            Describe = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Describe = DR("Describe").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Describe] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Title() As String
        Get
            Title = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Title = DR("Title").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Title] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Author() As String
        Get
            Author = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Author = DR("Author").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Author] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property BuildTime() As String
        Get
            BuildTime = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                BuildTime = DR("BuildTime").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [BuildTime] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property Adminer() As String
        Get
            Adminer = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Adminer = DR("Adminer").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Adminer] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
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
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
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
            CMD.CommandText = "UPDATE [Articles] SET [LastTime] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property FileName() As String
        Get
            FileName = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Articles] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                FileName = DR("FileName").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [FileName] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Shared Function GetArtNum() As Integer
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select count(*) from [Articles]"
        GetArtNum = Int(CMD.ExecuteScalar)
        CMD.Dispose()
        CN.Close()
    End Function
End Class
