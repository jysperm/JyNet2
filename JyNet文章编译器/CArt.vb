Imports Microsoft.VisualBasic
Imports System.Data.OleDb                                                          '引入Access数据库相关的命名空间
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
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [PV] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As Integer)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Type] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [KeyWord] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Describe] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Title] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Author] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [BuildTime] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [Adminer] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [LastTime] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
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
            CN.Close()
        End Get
        Set(ByVal value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "UPDATE [Articles] SET [FileName] = '" & value.ToString & "' WHERE [ID] = " & P_ID.ToString
            CMD.ExecuteNonQuery()
            CN.Close()
        End Set
    End Property
    Shared Function Build_DeleteTable() As Integer
        My.Application.DoEvents()
        Build_DeleteTable = 1
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "DELETE * FROM [Tmp_ArtBuild]"
        CMD.ExecuteNonQuery()
        CMD.Dispose()
        CN.Close()
    End Function
    Shared Function Build_IsInTable(ByVal ID As Integer) As Boolean
        My.Application.DoEvents()
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "SELECT * FROM [Tmp_ArtBuild] WHERE [AID] =" & ID
        Dim DR = CMD.ExecuteReader
        Build_IsInTable = DR.Read
        CMD.Dispose()
        DR.Close()
        CN.Close()
    End Function
    Shared Function Build_GetV(ByVal ID As Integer) As Integer
        My.Application.DoEvents()
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "SELECT * FROM [Tmp_ArtBuild] WHERE [AID] =" & ID
        Dim DR = CMD.ExecuteReader
        DR.Read()
        Build_GetV = DR("V")
        CMD.Dispose()
        DR.Close()
        CN.Close()
    End Function
    Shared Function Build_UpDateV(ByVal ID As Integer) As Integer
        My.Application.DoEvents()
        Build_UpDateV = 1
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "UPDATE [Tmp_ArtBuild] SET [V]=" & CArt.Build_GetV(ID) + 1 & " WHERE [AID] =" & ID
        CMD.ExecuteNonQuery()
        CMD.Dispose()
        CN.Close()
    End Function
    Shared Function Build_NewV(ByVal ID As Integer, ByVal V As Integer) As Integer
        My.Application.DoEvents()
        Build_NewV = 1
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "INSERT INTO [Tmp_ArtBuild] (AID, V) VALUES(" & ID & "," & V & ")"
        CMD.ExecuteNonQuery()
        CMD.Dispose()
        CN.Close()
    End Function
    Shared Function Build_OutTable() As String
        My.Application.DoEvents()
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "SELECT * FROM [Tmp_ArtBuild] ORDER BY [V] DESC"
        Dim DR = CMD.ExecuteReader
        Dim I = 1
        Dim tAbout = ""
        Do While DR.Read And I <= 10
            tAbout = tAbout & "<li><a target='_blank' href='http://" & CString.DoMain & "/" & New Tag(New CArt(DR("AID")).Type).PathName & "/" & New CArt(DR("AID")).FileName & ".htm'>" & New CArt(DR("AID")).Title & "</a></li>"
            I = I + 1
        Loop
        Build_OutTable = tAbout
        CMD.Dispose()
        DR.Close()
        CN.Close()
    End Function
    Shared Function Build(ByVal ID As Integer) As Integer
        Dim tOut As String = CFile.LoadFile("../UI/ArtTem.htm")
        Dim tArt As CArt = New CArt(ID)
        tOut = Replace(tOut, "{jybox~标题}", tArt.Title.ToString)
        Dim tTag As Tag = New Tag(tArt.Type)
        tOut = Replace(tOut, "{jybox~分类名}", tTag.Name.ToString)
        tOut = Replace(tOut, "{jybox~关键字}", tArt.KeyWord.ToString)
        tOut = Replace(tOut, "{jybox~描述}", tArt.Describe.ToString)
        tOut = Replace(tOut, "{jybox~分类关键字}", tTag.KeyWord.ToString)
        tOut = Replace(tOut, "{jybox~ID}", tArt.ID.ToString)
        tOut = Replace(tOut, "{jybox~菜单}", CMenu.GetHtml.ToString)
        tOut = Replace(tOut, "{jybox~作者}", tArt.Author.ToString)
        tOut = Replace(tOut, "{jybox~发表时间}", tArt.BuildTime.ToString)
        tOut = Replace(tOut, "{jybox~正文}", CFile.LoadFile("../App_Code/Art/" & ID.ToString & ".htm"))
        tOut = Replace(tOut, "{jybox~编辑}", tArt.Adminer.ToString)
        tOut = Replace(tOut, "{jybox~编辑时间}", tArt.LastTime.ToString)
        tOut = Replace(tOut, "{jybox~上一篇}", "<a href='http://" & CString.DoMain & "/" & New Tag(New CArt(ID - 1).Type).PathName & "/" & New CArt(ID - 1).FileName & ".htm'>" & New CArt(ID - 1).Title & "</a>")
        If ID = CArt.GetArtNum Then
            tOut = Replace(tOut, "{jybox~下一篇}", "没有了")
        Else
            tOut = Replace(tOut, "{jybox~下一篇}", "<a href='http://" & CString.DoMain & "/" & New Tag(New CArt(ID + 1).Type).PathName & "/" & New CArt(ID + 1).FileName & ".htm'>" & New CArt(ID + 1).Title & "</a>")
        End If
        '计算相关文章
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        Dim V As Integer
        CArt.Build_DeleteTable()
        CMD.Connection = CConnDB.ConnToDB(0)
        Dim DR As OleDbDataReader
        Dim K() = Split(tArt.KeyWord, ",")
        Dim I As Integer
        For I = 0 To K.Length - 1
            My.Application.DoEvents()
            CMD.CommandText = "SELECT * FROM [Articles] WHERE [KeyWord] LIKE '%" & K(I).ToString & "%'"
            DR = CMD.ExecuteReader
            Do While DR.Read
                If Build_IsInTable(DR("ID")) Then
                    CArt.Build_UpDateV(DR("ID"))
                Else
                    V = 0
                    If tArt.Type = DR("Type") Then
                        V = 3
                    End If
                    CArt.Build_NewV(DR("ID"), V)
                End If
            Loop
            DR.Close()
        Next
        tOut = Replace(tOut, "{jybox~相关文章}", CArt.Build_OutTable)
        CFile.WriteFile("../" & tTag.PathName & "/" & tArt.FileName & ".htm", tOut.ToString)
        Build = 1
    End Function
    Shared Function GetArtNum() As Integer
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select count(*) from [Articles]"
        GetArtNum = Int(CMD.ExecuteScalar)
    End Function
End Class
