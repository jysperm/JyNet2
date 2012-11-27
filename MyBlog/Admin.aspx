<%@ Page Language="VB"  ValidateRequest="False" CodeFile="Admin.aspx.vb" Inherits="MyBlog_Admin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <script src="../JS/KE.js" type="text/javascript"></script>
    <title>博客后台</title>
    <style type="text/css">
        body
        {
            margin: 120px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:TextBox ID="In" runat="server" TextMode="MultiLine" name="context" cols="100" rows="8" style="width: 700px; height: 300px;"></asp:TextBox>
        <script type="text/javascript">
            KE.show({
                id: 'In'
            });
        </script><br />
        标题<asp:TextBox ID="TTitle" runat="server" Width="694px"></asp:TextBox><br />
        关键字<asp:TextBox ID="KKey" runat="server" Width="694px"></asp:TextBox><br />
        确认码(防误点)，请输入jybox<asp:TextBox ID="TV" runat="server" Width="694px"></asp:TextBox>
        <br />
        &lt;link href="../CSS/HL.css" rel="Stylesheet" type="text/css">
        <br />
        <asp:Button ID="Button1" runat="server" Text="Button" Height="40px" 
            Width="100px" />
    </div>
    </form>
</body>
</html>
