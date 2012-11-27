Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
Public Class Tag
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
    Property Name() As String
        Get
            Name = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Art_Tag] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                Name = DR("TagName").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Art_Tag] SET [TagName] = '" & value.ToString & "' WHERE [ID] = '" & P_ID.ToString & "'"
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
            CMD.CommandText = "select * from [Art_Tag] where [ID]=" & P_ID.ToString
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
            CMD.CommandText = "UPDATE [Art_Tag] SET [KeyWord] = '" & value.ToString & "' WHERE [ID] = '" & P_ID.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Property PathName() As String
        Get
            PathName = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Art_Tag] where [ID]=" & P_ID.ToString
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then                                                        '如果有数据，读取
                PathName = DR("PathName").ToString
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Art_Tag] SET [PathName] = '" & value.ToString & "' WHERE [ID] = '" & P_ID.ToString & "'"
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
    Shared Function GetTagNum() As Integer
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select count(*) from [Art_Tag]"
        GetTagNum = Int(CMD.ExecuteScalar)
        CMD.Dispose()
        CN.Close()
    End Function
End Class
