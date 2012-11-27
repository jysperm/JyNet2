Imports System.Data.OleDb

Partial Class MyBlog_Index
    Inherits System.Web.UI.Page
    Public Function GetArt(ByVal PName As String, ByVal tAID As Integer) As String
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select * from [Blog_Art] where [ID]=" & tAID.ToString
        Dim DR As OleDbDataReader = CMD.ExecuteReader
        If DR.Read Then
        Else
            DR.Close()
            CMD.Dispose()
            CN.Close()
            Response.Redirect("Art.aspx")
        End If
        GetArt = DR(PName).ToString
        DR.Close()
        CMD.Dispose()
        CN.Close()

    End Function
End Class
