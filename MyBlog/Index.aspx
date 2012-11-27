<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Index.aspx.vb" Inherits="MyBlog_Index" %>

<%@ Import Namespace="System.Data.OleDb" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>精英王子's Blog 王子亭的个人博客</title>
    <meta name="keywords" content="精英┱┱王子,精英王子,精英盒子,王子亭,Jybox,Jingyingbox,精英,JY,王子,个人,博客,主页,Blog" />
    <meta name="description" content="这里是精英王子的博客，用来记录一些点滴，不写技术性文章" />
    <link href="../CSS/Frame.css" rel="stylesheet" type="text/css" />
    <link href="../CSS/WebDialog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebAjax.js" type="text/javascript"></script>
    <link href="../CSS/MyBlog.css" rel="stylesheet" type="text/css" />
    <script src="../JS/WebDialog.js" type="text/javascript"></script>
    <script src="../JS/WebCommon.js" type="text/javascript"></script>
    <script type="text/javascript">
        function ChineseCalendar(dateObj) {
            this.dateObj = (dateObj != undefined) ? dateObj : new Date();
            this.SY = this.dateObj.getFullYear();
            this.SM = this.dateObj.getMonth();
            this.SD = this.dateObj.getDate();
            this.lunarInfo = [0x04bd8, 0x04ae0, 0x0a570, 0x054d5, 0x0d260, 0x0d950, 0x16554, 0x056a0, 0x09ad0, 0x055d2,
            0x04ae0, 0x0a5b6, 0x0a4d0, 0x0d250, 0x1d255, 0x0b540, 0x0d6a0, 0x0ada2, 0x095b0, 0x14977,
            0x04970, 0x0a4b0, 0x0b4b5, 0x06a50, 0x06d40, 0x1ab54, 0x02b60, 0x09570, 0x052f2, 0x04970,
            0x06566, 0x0d4a0, 0x0ea50, 0x06e95, 0x05ad0, 0x02b60, 0x186e3, 0x092e0, 0x1c8d7, 0x0c950,
            0x0d4a0, 0x1d8a6, 0x0b550, 0x056a0, 0x1a5b4, 0x025d0, 0x092d0, 0x0d2b2, 0x0a950, 0x0b557,
            0x06ca0, 0x0b550, 0x15355, 0x04da0, 0x0a5b0, 0x14573, 0x052b0, 0x0a9a8, 0x0e950, 0x06aa0,
            0x0aea6, 0x0ab50, 0x04b60, 0x0aae4, 0x0a570, 0x05260, 0x0f263, 0x0d950, 0x05b57, 0x056a0,
            0x096d0, 0x04dd5, 0x04ad0, 0x0a4d0, 0x0d4d4, 0x0d250, 0x0d558, 0x0b540, 0x0b6a0, 0x195a6,
            0x095b0, 0x049b0, 0x0a974, 0x0a4b0, 0x0b27a, 0x06a50, 0x06d40, 0x0af46, 0x0ab60, 0x09570,
            0x04af5, 0x04970, 0x064b0, 0x074a3, 0x0ea50, 0x06b58, 0x055c0, 0x0ab60, 0x096d5, 0x092e0,
            0x0c960, 0x0d954, 0x0d4a0, 0x0da50, 0x07552, 0x056a0, 0x0abb7, 0x025d0, 0x092d0, 0x0cab5,
            0x0a950, 0x0b4a0, 0x0baa4, 0x0ad50, 0x055d9, 0x04ba0, 0x0a5b0, 0x15176, 0x052b0, 0x0a930,
            0x07954, 0x06aa0, 0x0ad50, 0x05b52, 0x04b60, 0x0a6e6, 0x0a4e0, 0x0d260, 0x0ea65, 0x0d530,
            0x05aa0, 0x076a3, 0x096d0, 0x04bd7, 0x04ad0, 0x0a4d0, 0x1d0b6, 0x0d250, 0x0d520, 0x0dd45,
            0x0b5a0, 0x056d0, 0x055b2, 0x049b0, 0x0a577, 0x0a4b0, 0x0aa50, 0x1b255, 0x06d20, 0x0ada0,
            0x14b63];

            //传回农历 y年闰哪个月 1-12 , 没闰传回 0
            this.leapMonth = function (y) {
                return this.lunarInfo[y - 1900] & 0xf;
            };
            //传回农历 y年m月的总天数
            this.monthDays = function (y, m) {
                return (this.lunarInfo[y - 1900] & (0x10000 >> m)) ? 30 : 29;
            };
            //传回农历 y年闰月的天数
            this.leapDays = function (y) {
                if (this.leapMonth(y)) {
                    return (this.lunarInfo[y - 1900] & 0x10000) ? 30 : 29;
                }
                else {
                    return 0;
                }
            };
            //传回农历 y年的总天数
            this.lYearDays = function (y) {
                var i, sum = 348;
                for (i = 0x8000; i > 0x8; i >>= 1) {
                    sum += (this.lunarInfo[y - 1900] & i) ? 1 : 0;
                }
                return sum + this.leapDays(y);
            };
            //算出农历, 传入日期对象, 传回农历日期对象
            //该对象属性有 .year .month .day .isLeap .yearCyl .dayCyl .monCyl
            this.Lunar = function (dateObj) {
                var i, leap = 0, temp = 0, lunarObj = {};
                var baseDate = new Date(1900, 0, 31);
                var offset = (dateObj - baseDate) / 86400000;
                lunarObj.dayCyl = offset + 40;
                lunarObj.monCyl = 14;
                for (i = 1900; i < 2050 && offset > 0; i++) {
                    temp = this.lYearDays(i);
                    offset -= temp;
                    lunarObj.monCyl += 12;
                }
                if (offset < 0) {
                    offset += temp;
                    i--;
                    lunarObj.monCyl -= 12;
                }

                lunarObj.year = i;
                lunarObj.yearCyl = i - 1864;
                leap = this.leapMonth(i);
                lunarObj.isLeap = false;
                for (i = 1; i < 13 && offset > 0; i++) {
                    if (leap > 0 && i == (leap + 1) && lunarObj.isLeap == false) {
                        --i;
                        lunarObj.isLeap = true;
                        temp = this.leapDays(lunarObj.year);
                    }
                    else {
                        temp = this.monthDays(lunarObj.year, i)
                    }
                    if (lunarObj.isLeap == true && i == (leap + 1)) {
                        lunarObj.isLeap = false;
                    }
                    offset -= temp;
                    if (lunarObj.isLeap == false) {
                        lunarObj.monCyl++;
                    }
                }

                if (offset == 0 && leap > 0 && i == leap + 1) {
                    if (lunarObj.isLeap) {
                        lunarObj.isLeap = false;
                    }
                    else {
                        lunarObj.isLeap = true;
                        --i;
                        --lunarObj.monCyl;
                    }
                }

                if (offset < 0) {
                    offset += temp;
                    --i;
                    --lunarObj.monCyl
                }
                lunarObj.month = i;
                lunarObj.day = offset + 1;
                return lunarObj;
            };
            //中文日期
            this.cDay = function (m, d) {
                var nStr1 = ['日', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];
                var nStr2 = ['初', '十', '廿', '卅', '　'];
                var s;
                if (m > 10) {
                    s = '十' + nStr1[m - 10];
                }
                else {
                    s = nStr1[m];
                }
                s += '月';
                switch (d) {
                    case 10:
                        s += '初十';
                        break;
                    case 20:
                        s += '二十';
                        break;
                    case 30:
                        s += '三十';
                        break;
                    default:
                        s += nStr2[Math.floor(d / 10)];
                        s += nStr1[d % 10];
                }
                return s;
            };
            this.solarDay2 = function () {
                var sDObj = new Date(this.SY, this.SM, this.SD);
                var lDObj = this.Lunar(sDObj);
                var tt = '农历' + this.cDay(lDObj.month, lDObj.day);
                lDObj = null;
                return tt;
            };
        }
        function Year_Month() {
            var now = new Date();
            var yy = now.getYear();
            var mm = now.getMonth();
            var mmm = new Array();
            mmm[0] = "1月";
            mmm[1] = "2月";
            mmm[2] = "3月";
            mmm[3] = "4月";
            mmm[4] = "5月";
            mmm[5] = "6月";
            mmm[6] = "7月";
            mmm[7] = "8月";
            mmm[8] = "9月";
            mmm[9] = "10月";
            mmm[10] = "11月";
            mmm[11] = "12月";
            mm = mmm[mm];
            return (mm);
        }

        function this_year() {
            var now = new Date();
            var yy = now.getYear();
            return (yy + 1900);
        }

        function Today() {
            var now = new Date();
            return (now.getDate());
        }

        function Day_of_Today() {
            var day = new Array();
            day[0] = "星期日";
            day[1] = "星期一";
            day[2] = "星期二";
            day[3] = "星期三";
            day[4] = "星期四";
            day[5] = "星期五";
            day[6] = "星期六";
            var now = new Date();
            return (day[now.getDay()]);
        }

        function CurentTime() {
            var now = new Date();
            var hh = now.getHours();
            var mm = now.getMinutes();
            var ss = now.getTime() % 60000;
            ss = (ss - (ss % 1000)) / 1000;
            var clock = hh + ':';
            if (mm < 10) clock += '0';
            clock += mm + ':';
            if (ss < 10) clock += '0';
            clock += ss;
            return (clock);
        }

        function refreshCalendarClock() {
            document.all.T2.innerHTML = Year_Month();
            document.all.T3.innerHTML = Today();
            document.all.T1.innerHTML = this_year();
            document.all.T4.innerHTML = CurentTime();
            document.all.T5.innerHTML = Day_of_Today();
            var date = new ChineseCalendar();
            document.all.T6.innerHTML = date.solarDay2();
        }
        setInterval('refreshCalendarClock()', 1000); 
    </script>
</head>
<body>
    <script type="text/javascript">
        var dialog = new WebDialog();
        dialog.ImgZIndex = 107;
        dialog.DialogZIndex = 108;
        dialog.Width = 400;
    </script>
    <form id="FormMain" runat="server">
    <div id="MainFrame">
        <div id="HeadArea">
            <div id="BMenu">
                <ul style="margin: 0px 0px 0px 0px">
                    <li><a href="http://jybox.net">网站主页</a></li>
                    <li><a href="http://jybox.net/MyBlog">博客主页</a></li>
                    <li><a href="http://jybox.net/MyBlog/Art.aspx">日志列表</a></li>
                    <li><a href="http://jybox.net/MyBlog/About.aspx">关于我</a></li>
                </ul>
            </div>
        </div>
        <div class="MainTable">
            <div id="ILeft">
                <div id="MyInfo">
                    <img alt="精英王子" class="Logo" src="../Images/MyLogo.gif" />
                    <div class="C">
                        精英王子</div>(精英┱┱王子)<br />
                    王子亭，男，15岁<br />
                    辽宁 沈阳
                    <div class="Clear">
                    </div>
                    邮箱：m@jybox.net(<a href="mailto:m@jybox.net">发送邮件</a>)<br />
                    QQ：172157947<a target="_blank" href="http://wpa.qq.com/msgrd?v=3&uin=172157947&site=qq&menu=yes"><img
                        src="http://wpa.qq.com/pa?p=2:172157947:41" alt="点击这里给我发消息" style="vertical-align: middle;"></a>
                </div>
                <div id="Time">
                    <span id="T1"></span>年<span id="T2"></span><span id="T3"></span>日<br />
                    <span id="T5"></span>
                    <br />
                    <span id="T4"></span>
                    <br />
                    <span id="T6"></span>
                </div>
            </div>
            <div id="IRight">
                <h3>
                    日志-最近更新</h3>
                <%
                    Dim CN As OleDbConnection = CConnDB.ConnToDB(0)
                    Dim CMD As New OleDbCommand
                    CMD.Connection = CN
                    CMD.CommandText = "select * from [Blog_Art] ORDER BY [ID] DESC"
                    Dim DR As OleDbDataReader = CMD.ExecuteReader
                    Dim I = 0
                    Do While DR.Read And I < +3
                %>
                <div class="ATitle">
                    <a href="http://<%=CString.DoMain%>/MyBlog/Art.aspx?id=<%=DR("ID").ToString%>">
                        <%= DR("Title").ToString%>
                    </a>
                </div>
                <div class="AInfo">
                    <%= DR("TheTime").ToString%>(<%= CString.MyTime(DR("TheTime").ToString)%>) | 访问量:
                    <%= DR("PV").ToString%>
                    PageView
                </div>
                <div class="ArtBody">
                    <%= Left(GetArt("Body", DR("ID")).ToString, 500)%>
                    <br /><br />
                    <div><a class="FRA" href="http://<%=CString.DoMain%>/MyBlog/Art.aspx?id=<%=DR("ID").ToString%>">
                        以上是摘要，点击这里查看全文</a></div>
                </div>
                <%
                    I = I + 1
                Loop
                DR.Close()
                CMD.Dispose()
                CN.Close()
                
                %>
            </div>
            <div class="Clear">
            </div>
        </div>
    </div>
    </form>
    <div id="F">
        <br />
        <br />
        <br />
        By 精英王子 2009-2011
    </div>
    <script type="text/javascript">
        var _bdhmProtocol = (("https:" == document.location.protocol) ? " https://" : " http://");
        document.write(unescape("%3Cscript src='" + _bdhmProtocol + "hm.baidu.com/h.js%3F77016691cd5a049005dba568b5164b59' type='text/javascript'%3E%3C/script%3E"));
    </script>
</body>
</html>
