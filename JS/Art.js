function insertAtCursor(myField, myValue) {
  if (document.selection) {
    myField.focus();
    sel = document.selection.createRange();
    sel.text = myValue;
    sel.select();
  }
  else if (myField.selectionStart || myField.selectionStart == '0') {
    var startPos = myField.selectionStart;
    var endPos = myField.selectionEnd;
    var restoreTop = myField.scrollTop;
    myField.value = myField.value.substring(0, startPos) + myValue + myField.value.substring(endPos, myField.value.length);
    if (restoreTop > 0) {
      myField.scrollTop = restoreTop;
    }
    myField.focus();
    myField.selectionStart = startPos + myValue.length;
    myField.selectionEnd = startPos + myValue.length;
  } else {
    myField.value += myValue;
    myField.focus();
  }
}
function AddEm(n)
{
    insertAtCursor($("In"), "[/表情"+n+"]");
}
function ReVCode()
{
    times++;
    document.getElementById("VCodeImg").src="";
    document.getElementById("VCodeImg").src="../UI/VCode.aspx?times=" + times ;
}

function ALoad()
{
    $("PV").innerHTML = executeHttpRequest("POST", "../AJAX/Art_PV.ashx?mode=post", "id=" + id);
    $("RPV").innerHTML = $("PV").innerHTML;
    var result = executeHttpRequest("POST", "../AJAX/Art_Say.ashx?mode=post","do=getlogin");
    Show(10);
    if(result=="null")
    {
        $("Info").innerHTML="<span style='color:red;'>[您无需登录也可以发表评论!]</span>&nbsp;&nbsp;<a href='#' onclick='Login(ALoad);' style='font-size: 13px'>[点击这里登录(无需刷新)]</a>";
    }
    else
    {
        $("Info").innerHTML="["+result+"]写下你的看法吧！";
        $("VCodeIn").style.display="none";
    }
}
var times=1;
ALoad();
function NewSay()
{
    if($("In").value=="")
    {
        alert("请输入评论内容");
        return;
    }
    result = executeHttpRequest("POST", "../AJAX/Art_Say.ashx?mode=post","do=new&aid="+id+"&body="+encodeURIComponent($("In").value)+"&vcode="+$("VCode").value);
    ReVCode();
    if(result=="OK")
    {
        alert("评论成功");
        $("In").value="";
        Show(10);
    }
    else
    {
        alert(result);
    }
}
function Show(n)
{
    $("Say").innerHTML = executeHttpRequest("POST", "../AJAX/Art_Say.ashx?mode=post","do=get&num="+n+"&aid="+id);
}