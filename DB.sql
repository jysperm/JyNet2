CREATE TABLE [Art_Key]([ID] LONG NOT NULL,[Title] TEXT(255),[Type] LONG,[PV] LONG DEFAULT 0);
CREATE TABLE [Art_Say]([ID] LONG NOT NULL,[IsLogin] LONG NOT NULL,[Body] MEMO NOT NULL,[UName] TEXT(255) NOT NULL,[IP] TEXT(255) NOT NULL,[TheTime] DATETIME NOT NULL DEFAULT =Now(),[AID] LONG);
CREATE TABLE [Art_Tag]([ID] LONG NOT NULL,[TagName] TEXT(255) NOT NULL,[KeyWord] TEXT(255) NOT NULL,[PathName] TEXT(255) NOT NULL);
CREATE TABLE [Articles]([ID] LONG NOT NULL,[Type] LONG NOT NULL,[Title] TEXT(255) NOT NULL,[KeyWord] TEXT(255) NOT NULL,[Describe] TEXT(255) NOT NULL DEFAULT ����,[BuildTime] DATETIME,[LastTime] DATETIME,[Author] TEXT(255) NOT NULL,[Adminer] TEXT(255) NOT NULL,[PV] LONG DEFAULT 0,[FileName] TEXT(255),[Res] LONG DEFAULT 0,[LastRe] DATETIME);
CREATE TABLE [BBS_B]([ID] COUNTER NOT NULL,[BName] TEXT(255),[ReadMe] TEXT(255),[Logo] TEXT(255),[NumRe] LONG,[NumPV] LONG,[LastPID] LONG,[uMoney] LONG,[uExp] LONG,[uLv] LONG);
CREATE TABLE [BBS_P]([ID] LONG NOT NULL,[BID] LONG,[UName] TEXT(255),[tTime] DATETIME,[LastTime] DATETIME,[IP] TEXT(255),[Body] MEMO,[Title] TEXT(255),[Link] TEXT(255),[UpSet] LONG,[PV] LONG);
CREATE TABLE [BBS_R]([ID] LONG NOT NULL,[PID] LONG,[UName] TEXT(255),[tTime] DATETIME,[IP] TEXT(255),[Body] MEMO);
CREATE TABLE [Blog_Art]([ID] COUNTER NOT NULL,[Title] TEXT(255),[KeyWords] TEXT(255),[Body] MEMO,[TheTime] DATETIME,[PV] INTEGER);
CREATE TABLE [Blog_Say]([ID] COUNTER NOT NULL,[IsLogin] INTEGER,[Body] MEMO,[UName] TEXT(255),[IP] TEXT(255),[TheTime] DATETIME DEFAULT =Now(),[AID] INTEGER);
CREATE TABLE [Net_KY]([ID] COUNTER NOT NULL,[Key] TEXT(255),[TValue] TEXT(255),[����] TEXT(255));
CREATE TABLE [Net_User]([ID] LONG NOT NULL,[UserName] TEXT(255) NOT NULL,[NName] TEXT(255) NOT NULL,[PWD] TEXT(255) NOT NULL,[VIP] LONG DEFAULT 0,[EXP] LONG DEFAULT 10,[Money] LONG DEFAULT 10,[Gold] LONG DEFAULT 0,[RegTime] DATETIME DEFAULT =Now(),[QQ] TEXT(255),[Logo] MEMO DEFAULT "http://jybox.net/images/nologo.gif",[EMail] TEXT(255) NOT NULL,[RegIP] TEXT(255) NOT NULL,[LastIP] TEXT(255) NOT NULL,[LastTime] DATETIME DEFAULT =Now(),[Lv] LONG DEFAULT 0,[ConDay] LONG DEFAULT 0,[LastCard] DATETIME DEFAULT =Now());
CREATE TABLE [Tmp_ArtBuild]([ID] COUNTER NOT NULL,[AID] INTEGER,[V] INTEGER);