

var dialog = new WebDialog();
document.body.oncopy=function(){setTimeout(function(){var text=clipboardData.getData("text");if(text){text=text+"\r\n本文来自: 精英盒子.Net(jybox.net) ，欢迎来访！详细出处请参考："+location.href;clipboardData.setData("text",text)}},100)}

//统计代码
function left(mainStr,lngLen) { if (lngLen>0) {return mainStr.substring(0,lngLen)} else{return null} } 
function right(mainStr,lngLen) { if (mainStr.length-lngLen>=0 && mainStr.length>=0 && mainStr.length-lngLen<=mainStr.length) { return mainStr.substring(mainStr.length-lngLen,mainStr.length)} else{return null} } 
function mid(mainStr,starnum,endnum){ if (mainStr.length>=0){ return mainStr.substr(starnum,endnum) }else{return null} } 

function Root(u)
{
    var p=location.pathname;
    for(var r="",i=1,t=0;i<p.length;i++,t++)
        if(mid(p,i,1)=="/" && t>1)
            r=r+"../";
    return r+u
}
$("Footer").innerHTML=executeHttpRequest("POST", Root("AJAX/Footer.ashx?mode=post"),"do=get");

function MShow(n)
{
    $("FloatMenu").innerHTML=executeHttpRequest("POST", Root("UI/FloatMenu.aspx?mode=post"),"id=" + n);
}
MShow(100);



function correctPNG() {
    for (var i = 0; i < document.images.length; i++) {
        var img = document.images[i];
        var imgName = img.src.toUpperCase();
        if (imgName.substring(imgName.length - 3, imgName.length) == "PNG") {
            var imgID = (img.id) ? "id='" + img.id + "' " : "";
            var imgClass = (img.className) ? "class='" + img.className + "' " : "";
            var imgTitle = (img.title) ? "title='" + img.title + "' " : "title='" + img.alt + "' ";
            var imgStyle = "display:inline-block;" + img.style.cssText;
            if (img.align == "left") imgStyle = "float:left;" + imgStyle;
            if (img.align == "right") imgStyle = "float:right;" + imgStyle;
            if (img.parentElement.href) imgStyle = "cursor:hand;" + imgStyle;
            var strNewHTML = "<span " + imgID + imgClass + imgTitle + "style=\"" + "width:" + img.width + "px; height:" + img.height + "px;" + imgStyle + ";"
+ "filter:progid:DXImageTransform.Microsoft.AlphaImageLoader" + "(src='" + img.src + "', sizingMethod='scale');\"></span>";
            img.outerHTML = strNewHTML;
            i = i - 1;
        }
    }
}
window.attachEvent("onload", correctPNG);