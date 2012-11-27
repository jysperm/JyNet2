
Partial Class _My
    Inherits System.Web.UI.Page
    Public TU As CUser, IsLoged As Integer = 0
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        IsLoged = CLogin.Login()
        TU = New CUser(CLogin.GetUName)
    End Sub
End Class
