Imports Microsoft.VisualBasic
Imports System.Data.OleDb
Public Class CBBSRe
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
    Public Function GetReNum() As Integer
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select count(*) from [BBS_R] WHERE [PID]=" & P_ID.ToString
        GetReNum = Int(CMD.ExecuteScalar)
        CMD.Dispose()
        CN.Close()
    End Function
End Class
