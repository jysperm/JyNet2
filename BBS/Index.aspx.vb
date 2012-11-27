Imports System.Data.OleDb

Partial Class BBS_Index
    Inherits System.Web.UI.Page
    Public CN As OleDbConnection = CConnDB.ConnToDB(0)
    Public CMD As New OleDbCommand()
    Public TID As Integer = -1
    Public ListPage As Integer
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        CMD.Connection = CN
        If Not Request.QueryString("id") = Nothing Then
            If IsNumeric(Request.QueryString("id").ToString) Then
                TID = Request.QueryString("id").ToString
                If Not Request.QueryString("page") = Nothing Then
                    If Not IsNumeric(Request.QueryString("page").ToString) Then
                        ListPage = 1
                    Else
                        ListPage = Int(Request.QueryString("page").ToString)
                    End If
                Else
                    ListPage = 1
                End If
            Else
                Response.Redirect("Index.aspx")
            End If
        End If
    End Sub
    Public Function GetB(ByVal PName As String, Optional ByVal tTID As Integer = Nothing) As String
        Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
        Dim CMD As New OleDbCommand
        tTID = IIf(tTID = Nothing, TID, tTID)
        CMD.Connection = CN
        CMD.CommandText = "select * from [BBS_B] where [ID]=" & tTID.ToString
        Dim DR As OleDbDataReader = CMD.ExecuteReader
        If DR.Read Then
        Else
            DR.Close()
            CMD.Dispose()
            CN.Close()
            Response.Redirect("Index.aspx")
        End If
        GetB = DR(PName).ToString
        DR.Close()
        CMD.Dispose()
        CN.Close()
    End Function
    Public Function GetTitle() As String
        GetTitle = ""
        If TID <> -1 Then
            GetTitle = GetB("BName").ToString() & "_"
        End If
    End Function
    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        CMD.Dispose()
        CN.Close()
    End Sub
End Class
