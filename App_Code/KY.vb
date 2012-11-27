Imports Microsoft.VisualBasic
Imports System.Data.OleDb

Public Class KY
    Public Shared Property K(ByVal Index As String) As String
        Get
            K = ""
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_KY] where [Key]='" & Index.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then
                K = DR("TValue").ToString
            Else
                K = ""
            End If
            DR.Close()
            CMD.Dispose()
            CN.Close()
        End Get
        Set(ByVal Value As String)
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "select * from [Net_KY] where [Key]='" & Index.ToString & "'"
            Dim DR As OleDbDataReader = CMD.ExecuteReader
            If DR.Read Then
                Dim CMD2 As New OleDbCommand
                CMD2.Connection = CN
                CMD2.CommandText = "UPDATE [Net_KY] SET [TValue] = '" & Value.ToString & "' WHERE [Key] = '" & Index.ToString & "'"
                CMD2.ExecuteNonQuery()
                CMD2.Dispose()
            Else
                Dim CMD2 As New OleDbCommand
                CMD2.Connection = CN
                CMD2.CommandText = "INSERT INTO [Net_KY] (Key, TValue) VALUES('" & Index.ToString & "','" & Value.ToString & "')"
                CMD2.ExecuteNonQuery()
                CMD2.Dispose()
            End If
            CMD.Dispose()
            CN.Close()
        End Set
    End Property
End Class
