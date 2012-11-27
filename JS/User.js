
function Login(f)
{
    //参数f是登录后要执行的回调函数
    dialog.ImgZIndex = 107;
    dialog.DialogZIndex = 108;
    dialog.Width = 300;
    dialog.Text='登录';
    dialog.Content =executeHttpRequest("GET", Root("UI/JSLogin.htm"),null);
    dialog.OK=function()
    {
        OKClick(f);
    }
    dialog.Show(1);
}
function UInfo(UName)
{
    
}
function OKClick(f)
{
    var RV = /^[^<>,\n]{4,16}$/
    if(!RV.test($("TextU").value))
    {
        alert("请输入正确的用户名");
        return;
    }
    RV = /^\w{6,16}$/
    if(!RV.test($("TextP").value))
    {
        alert("请输入正确的密码");
        return;
    }
    var result = executeHttpRequest("POST", Root("AJAX/Login.ashx?mode=post"),"do=login&user=" + encodeURIComponent(document.getElementById("TextU").value) + "&pwd=" + encodeURIComponent(document.getElementById("TextP").value) + "&vcode=0")
    if(result=="OK")
    {   
        dialog.Close();
        alert("登录成功");
        f();
        return;
    }
    else if(result=="err")
    {
        alert("用户名或密码错误，请重新登录");
        return;
    }
    else if(result=="vcodeerr")
    {
        alert("验证码错误，请重新输入验证码");
        return;
    }
    else
    {
        alert("遇到不明错误。如频繁出现此错误请联系管理员(m@jybox.net)");
        return;
    }
}