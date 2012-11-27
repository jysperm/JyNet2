function $(divID)
{
    return document.getElementById(divID);
}
var createImg = function()
{
    return document.createElement("img");
}
var createDiv = function()
{
    return document.createElement("div");
}
var createBtn = function()
{
    var btn = document.createElement("input");
    btn.type = "button";
    return btn;
}
var createSpan = function()
{
    return document.createElement("span");
}
