<%@ Page Language="VB" %>

<%@ Import Namespace="System.Drawing" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Select Case Request.QueryString("do")
            Case "up"
                '如果有一个文件
                If Request.Files.Count = 1 Then
                    '取第一个文件
                    Dim File As HttpPostedFile = Request.Files(0)
                    '判断是否登录
                    If CLogin.Login() Then
                        '取扩展名
                        Dim MName = UCase(Right(File.FileName, 3))
                        '判断扩展名
                        If MName = "JPG" Or MName = "BMP" Or MName = "GIF" Or MName = "PNG" Or MName = "ICO" Then
                            '判断文件大小，并扣减积分
                            If File.ContentLength > 30 * 1024 Then
                                Dim TU = New CUser(CLogin.GetUName)
                                If TU.Money > ((File.ContentLength - 30 * 1024) / 1024) * 0.1 Then
                                    TU.Money -= ((File.ContentLength - 30 * 1024) / 1024) * 0.1
                                Else
                                    Response.Write("<span id='re'>用户金币不足，请选择小一些的图片</span>")
                                    Response.End()
                                End If
                            End If
                            '取顺序文件名
                            Dim FileName = CFile.XMLKey(CFile.LoadFile(Server.MapPath("~/Admin/Setting.ini")), "UImgs")
                            '保存文件
                            File.SaveAs(Server.MapPath("~/U_Imgs/") & FileName & "." & MName)
                            '更新顺序文件名
                            CFile.WriteFile(Server.MapPath("~/Admin/Setting.ini"), CFile.XMLUp(CFile.LoadFile(Server.MapPath("~/Admin/Setting.ini")), "UImgs", Int(FileName) + 1))
                            
                            '下面四行用于把物理路径转化成虚拟路径
                            Dim RootPath = Server.MapPath("~/")
                            Dim Path = Server.MapPath("~/U_Imgs/") & FileName & "." & MName
                            Path = Path.Remove(0, RootPath.Length)
                            Path = Path.Replace("\", "/")
                            
                            Response.Write("<span id='re'>OK</span><span id='Url'>" & Path & "</span>")
                        Else
                            Response.Write("<span id='re'>不支持的文件类型</span>")
                        End If
                    Else
                        Response.Write("<span id='re'>未登录或登录状态失效</span>")
                    End If
                Else
                    Response.Write("<span id='re'>没有选择文件</span>")
                End If
            Case Else
                Response.Write("<span id='re'>错误</span>")
        End Select
    End Sub
</script>

<script type="text/javascript">
document.getElementById("re").style.display ="none";
document.getElementById("Url").style.display ="none";
if(document.getElementById("re").innerHTML=="OK")
{
    window.parent.EndUp();
}
</script>

<form id="FromPC" method="post" runat="server" action="Img_Up.aspx?do=up" enctype="multipart/form-data">
上传文件：<input id="FileUp" runat="server" type="file" onclick="window.parent.document.getElementById('OpFromPC').checked=true" />
<br />
仅允许上传JPG、BMP、GIF、PNG、ICO格式的图像
<br />
大小若超过30KB，每KB收取0.1金币
</form>
