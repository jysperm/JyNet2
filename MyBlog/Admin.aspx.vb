Imports System.Data.OleDb

Partial Class MyBlog_Admin
    Inherits System.Web.UI.Page
    Protected Sub form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles form1.Load
        If CLogin.Login() Then
            Dim TU As CUser = New CUser(CLogin.GetUName)
            If TU.Lv >= 99 Then
                Exit Sub
            Else
                Response.Write("错误.您不是管理员")
                Response.End()
            End If
        Else
            Response.Write("错误.未登录或登录状态失效")
            Response.End()
        End If
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TV.Text = "jybox" Then
            Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
            Dim CMD As New OleDbCommand
            CMD.Connection = CN
            CMD.CommandText = "INSERT INTO [Blog_Art] (Title, KeyWords, Body,TheTime,PV) VALUES(@Title, @KeyWords, @Body,@TheTime,@PV)"
            '向SQL语句中补充数据
            CMD.Parameters.AddWithValue("@Title", TTitle.Text)
            CMD.Parameters.AddWithValue("@KeyWords", KKey.Text)
            CMD.Parameters.AddWithValue("@Body", [In].Text)         '对密码MD5后存入数据库
            CMD.Parameters.AddWithValue("@TheTime", Now.ToString)
            CMD.Parameters.AddWithValue("@PV", 0)
            '执行命令
            CMD.ExecuteNonQuery()
            CMD.Dispose()
            CN.Close()
            Response.Write("成功")
        Else
            Response.Write("防误点！")
        End If
    End Sub
End Class
