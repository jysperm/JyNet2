Imports System.Data.OleDb

Partial Class MyBlog_Art
    Inherits System.Web.UI.Page
    '0为List，1为Art
    Public ArtList As Integer = 0
    Public AtrTitle As String
    Public AID As Integer
    Public ListPage As Integer
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.QueryString("id") = Nothing Then
            '文章列表
            ArtList = 0
            If Not Request.QueryString("id") = Nothing Then
                If Not IsNumeric(Request.QueryString("page").ToString) Then
                    ListPage = 1
                Else
                    ListPage = Int(Request.QueryString("page").ToString)
                End If
            Else
                ListPage = 1
            End If
        Else
            If Not IsNumeric(Request.QueryString("id").ToString) Then
                Response.Redirect("Art.aspx")
            End If
            '文章页
            ArtList = 1
            AID = Int(Request.QueryString("id").ToString)
            ArtPV()
        End If

    End Sub
    Public Function GetArt(ByVal PName As String, Optional ByVal tAID As Integer = Nothing) As String
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        tAID = IIf(tAID = Nothing, AID, tAID)
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
    Public Function GetArts() As Integer
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "select count(*) from [Blog_Art]"
        GetArts = Int(CMD.ExecuteScalar)
        CMD.Dispose()
        CN.Close()
    End Function
    Public Sub ArtPV()
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        CMD.Connection = CN
        CMD.CommandText = "UPDATE [Blog_Art] SET [PV] = [PV]+1 WHERE ID=" & AID.ToString
        CMD.ExecuteNonQuery()
        CMD.Dispose()
        CN.Close()
    End Sub
End Class
