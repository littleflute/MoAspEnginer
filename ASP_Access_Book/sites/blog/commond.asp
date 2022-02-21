<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit
Response.Buffer = True
Server.ScriptTimeOut = 90

Session.CodePage=65001

If Trim(Request.QueryString("CP"))="GBK" Then Session.CodePage = 936

'定义 Cookie,Application 域，必须修改，否则可能运行不正常
Const CookieName="acblog"

'站点开关操作
IF Not isNumeric(Application(CookieName & "_SiteEnable")) Then
	Application.Lock
	Application(CookieName & "_SiteEnable") = 1
	Application(CookieName & "_SiteDisbleWhy") = ""
	Application.UnLock
End IF
IF Application(CookieName & "_SiteEnable") = 0 AND Application(CookieName & "_SiteDisbleWhy")<>"" AND inStr(Replace(Lcase(Request.ServerVariables("URL")),"\","/"),"/admincp.asp") = 0  AND inStr(Replace(Lcase(Request.ServerVariables("URL")),"\","/"),"/logging.asp") = 0 Then
	Response.Write(Application(CookieName & "_SiteDisbleWhy"))
	Response.End
End IF

Dim StartTime,SQLQueryNums
StartTime=Timer()
SQLQueryNums=0

DIM MaxUrl,MaxHttp
IF session("memStatus")<>"Admin" AND session("memStatus")<>"SupAdmin" Then
   MaxUrl=2
   MaxHttp=4
else
   MaxUrl=50
   MaxHttp=50
end if

'定义数据库链接文件，根据自己的情况修改
Const AccessPath="51pcbook"
Const AccessFile="blog.mdb" 
Const IPAccessFile="ipdata.asa"
dim db,pass_word,user_id,data_source,connstr
'定义数据库连接
Dim Conn
On Error Resume Next
Set Conn= Server.CreateObject("ADODB.Connection")
	connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&AccessFile&"")
	conn.Open connstr
If Err Then
    Err.Clear
    Set Conn = Nothing
    Response.Write("<head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title>数据库连接出错，请检查连接字串</title></head><body><div align=""center"" style=""width:400px;height:100px;padding: 8px;font-size:9pt;border: 1px solid ThreeDShadow;POSITION:absolute;top:expression((document.body.offsetHeight-100)/2);left:expression((document.body.offsetWidth-400)/2);""><table width=""100%"" height=""100%"" style=""font-size:12px;font-family:Tahoma;""><tr><td align=""center""><strong>数据库连接出错，请检查连接字串</strong></td></tr></table></div></body>")
    Response.End
End If

dim Badwords
BadWords="*****"'这里的*为你要过滤的字

function ChkBadWords(strContent)
    dim bwords,i
    bwords = split(BadWords, "|")
    for i = 0 to ubound(bwords)
        strContent = Replace(strContent, bwords(i), string(len(bwords(i)),"×")) 
    next
    strContent = Replace(strContent, "'","") '顺便过滤了一下单引号
    ChkBadWords = strContent
end function

Dim SQL,TempVar,siteTitle '定义常用变量

Dim blog_Infos,SiteName,SiteUrl,blogPerPage,blog_LogNums,blog_CommNums,blog_MemNums,blog_VisitBaseNums,blog_VisitNums,blog_QuoteNums,blog_GuestbookNums,blog_downloadNums,blog_SpamIP,blog_photoNums,threadPerPage,postPerPage,blog_New,blog_PostNums,blog_ThreadNums
Set blog_Infos=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM blog_Info"
blog_Infos.Open SQL,Conn,1,1
SQLQueryNums=SQLQueryNums+1
If blog_Infos.EOF And blog_Infos.BOF Then
	Response.Write("站点出错，请检查数据库中的站点基本信息设置……")
	Response.End
Else
	SiteName=blog_Infos("blog_Name")
	SiteURL=blog_Infos("blog_URL")
	blogPerPage=blog_Infos("blog_PerPage")
	blog_LogNums=blog_Infos("blog_LogNums")
	blog_CommNums=blog_Infos("blog_CommNums")
	blog_GuestbookNums=blog_Infos("blog_GuestbookNums")
	blog_MemNums=blog_Infos("blog_MemNums")
	blog_QuoteNums=blog_Infos("blog_QuoteNums")
	blog_VisitBaseNums=blog_Infos("blog_VisitBaseNums")
	blog_VisitNums=blog_Infos("blog_VisitNums")+blog_VisitBaseNums
	blog_downloadNums=blog_Infos("blog_downloadNums")
	blog_SpamIP=blog_Infos("blog_SpamIP")
	blog_photoNums=blog_Infos("blog_photoNums")
	blog_ThreadNums=blog_Infos("blog_ThreadNums")
	blog_PostNums=blog_Infos("blog_PostNums")
	threadPerPage=blog_Infos("thread_PerPage")
	postPerPage=blog_Infos("post_PerPage")
	blog_New=blog_Infos("blog_New")
End If
blog_Infos.Close
Set blog_Infos=Nothing

Dim Arr_Download
IF Not IsArray(Application(CookieName&"_blog_Download")) Then
	Dim log_Download,log_DownloadList
	Set log_Download=Conn.Execute("SELECT * FROM blog_Download")
	TempVar=""
	Do While Not log_Download.EOF
		log_DownloadList=log_DownloadList&TempVar&log_Download("downl_ID")&"$|$"&log_Download("downl_Cate")&"$|$"&log_Download("downl_Name")&"$|$"&log_Download("downl_Author")&"$|$"&log_Download("downl_From")&"$|$"&log_Download("downl_FromURL")&"$|$"&log_Download("downl_ImgPath")&"$|$"&log_Download("downl_PostTime")&"$|$"&log_Download("downl_Comment")&"$|$"&log_Download("downl_Dcomm1")&"$|$"&log_Download("downl_Dcomm1URL")&"$|$"&log_Download("downl_Dcomm2")&"$|$"&log_Download("downl_Dcomm2URL")&"$|$"&log_Download("downl_Dcomm3")&"$|$"&log_Download("downl_Dcomm3URL")&"$|$"&log_Download("downl_Dcomm4")&"$|$"&log_Download("downl_Dcomm4URL")&"$|$"&log_Download("downl_Nums")
		TempVar="$,$"
		log_Download.MoveNext
	Loop
	Set log_Download=Nothing
	Arr_Download=Split(log_DownloadList,"$,$")
	Application.Lock
	Application(CookieName&"_blog_Download")=Arr_Download
	Application.UnLock
Else
	Arr_Download=Application(CookieName&"_blog_Download")
End IF

Dim Guest_IP
Guest_IP=Replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"'","")
If Guest_IP=Empty Then Guest_IP=Replace(Request.ServerVariables("REMOTE_ADDR"),"'","")

Dim SpamIP,SpamIPEvery
SpamIP=Split(blog_SpamIP,";")
For Each SpamIPEvery In SpamIP
	If Guest_IP = SpamIPEvery Then
		Response.Redirect("http://www.mii.gov.cn/mii/")
		Response.Write.End
	End If
Next

'站点统计代码
If Session("GuestIP")<>Guest_IP Then
	Dim Guest_Agent,Guest_Month,Guest_Week,Guest_Hour,Guest_OS,Guest_Browser
	Guest_Agent=Trim(Request.ServerVariables("HTTP_USER_AGENT"))
	Guest_Month=Month(Now())
	If Len(Guest_Month)<2 Then 
		Guest_Month=Year(Now())&"0"&Guest_Month
	Else
		Guest_Month=Year(Now())&Guest_Month
	End If
	Guest_Week=WeekDay(Now())-1
	Guest_Hour=Hour(Now())
	If Len(Guest_Hour)<2 Then Guest_Hour="0"&Guest_Hour
	If InStr(Guest_Agent,"Win") Then
		Guest_OS="Windows"
	ElseIf InStr(Guest_Agent,"Mac") Then
		Guest_OS="Mac"
	ElseIf InStr(Guest_Agent,"Linux") Then
		Guest_OS="Linux"
	ElseIf InStr(Guest_Agent,"FreeBSD") Then
		Guest_OS="FreeBSD"
	ElseIf InStr(Guest_Agent,"SunOS") Then
		Guest_OS="SunOS"
	ElseIf InStr(Guest_Agent,"BeOS") Then
		Guest_OS="BeOS"
	ElseIf InStr(Guest_Agent,"OS/2") Then
		Guest_OS="OS/2"
	ElseIf InStr(Guest_Agent,"AIX") Then
		Guest_OS="AIX"
	ElseIf InStr(Guest_Agent,"search") Or InStr(Guest_Agent,"Spider") Or InStr(Guest_Agent,"Googlebot") Then
		Guest_OS="Search"
	Else
		Guest_OS="Other"
	End If
	If InStr(Guest_Agent,"Maxthon") Or InStr(Guest_Agent,"MyIE") Then
		Guest_Browser="Maxthon"
	ElseIf InStr(Guest_Agent,"MSIE") Then
		Guest_Browser="MSIE"
	ElseIf InStr(Guest_Agent,"Netscape") Then
		Guest_Browser="Netscape"
	ElseIf InStr(Guest_Agent,"Konqueror") Then
		Guest_Browser="Konqueror"
	ElseIf InStr(Guest_Agent,"Firefox") Then
		Guest_Browser="Firefox"
	ElseIf InStr(Guest_Agent,"search") Or InStr(Guest_Agent,"Spider") Or InStr(Guest_Agent,"Googlebot") Then
		Guest_Browser="Search"
	ElseIf InStr(Guest_Agent,"Reader") Or InStr(Guest_Agent,"FeedDemon") Then
		Guest_Browser="RSSReader"
	Else
		Guest_Browser="Other"
	End If
	If Conn.ExeCute("SELECT COUNT(coun_Char) FROM blog_Counter WHERE coun_Type='Month' AND coun_Char='"&Guest_Month&"'")(0)=0 Then
		Conn.ExeCute("INSERT INTO blog_Counter(coun_Type,coun_Char,coun_Nums) VALUES ('Month','"&Guest_Month&"',0)")
		SQLQueryNums=SQLQueryNums+1
	End If
	Conn.ExeCute("UPDATE blog_Counter SET coun_Nums=coun_Nums+1 WHERE (coun_Type='OS' AND coun_Char='"&Guest_OS&"') OR (coun_Type='Browser' AND coun_Char='"&Guest_Browser&"') OR (coun_Type='Week' AND coun_Char='"&Guest_Week&"') OR (coun_Type='Hour' AND coun_Char='"&Guest_Hour&"') OR (coun_Type='Month' AND coun_Char='"&Guest_Month&"')")
	Conn.ExeCute("UPDATE blog_Info SET blog_VisitNums=blog_VisitNums+1")
	SQLQueryNums=SQLQueryNums+3
	Session("GuestIP")=Guest_IP
End If

Dim memName,memPassword,memStatus
memName=CheckStr(Request.Cookies(CookieName)("memName"))
memPassword=CheckStr(Request.Cookies(CookieName)("memPassword"))
memStatus=CheckStr(Request.Cookies(CookieName)("memStatus"))

IF memName<>Empty Then
	Dim CheckCookie
	Set CheckCookie=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT mem_Name,mem_Password,mem_Status,mem_LastIP FROM blog_Member WHERE mem_Name='"&memName&"' AND mem_Password='"&memPassword&"' AND mem_Status='"&memStatus&"'"
	CheckCookie.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If CheckCookie.EOF AND CheckCookie.BOF Then
		Response.Cookies(CookieName)("memName")=""
		memName=Empty
		Response.Cookies(CookieName)("memPassword")=""
		memPassword=Empty
		Response.Cookies(CookieName)("memStatus")=""
		memStatus=Empty
	Else
		If CheckCookie("mem_LastIP")<>Guest_IP Or isNull(CheckCookie("mem_LastIP")) Then
			Response.Cookies(CookieName)("memName")=""
			memName=Empty
			Response.Cookies(CookieName)("memPassword")=""
			memPassword=Empty
			Response.Cookies(CookieName)("memStatus")=""
			memStatus=Empty
		End If
	End IF
	CheckCookie.Close
	Set CheckCookie=Nothing
Else
	Response.Cookies(CookieName)("memName")=""
	memName=Empty
	Response.Cookies(CookieName)("memPassword")=""
	memPassword=Empty
	Response.Cookies(CookieName)("memStatus")=""
	memStatus=Empty
End IF

'上传文件的大小以及后缀名限制
Dim Adm_UP_FileSize,Adm_UP_FileType,Mem_UP_FileSize,Mem_UP_FileType,MemCanUP
MemCanUP=0             '设定一般用户是否可以上传文件，1为可以上传，0为不可以上传
Adm_UP_FileSize = 20480000
Adm_UP_FileType = "RAR,ZIP,SWF,JPG,PNG,GIF,DOC,TXT,CHM,PDF,ACE,JPG,MP3,WMA,WMV,MIDI,AVI,RM,RA,RMVB,MOV,TORRENT"
Mem_UP_FileSize = 1024000
Mem_UP_FileType = "RAR,ZIP,SWF,JPG,PNG,GIF,DOC,TXT,CHM,PDF,ACE,JPG,MP3,WMA,WMV,MIDI,AVI,RM,RA,RMVB,MOV,TORRENT"


'写入日志分类
Dim Arr_Category
IF Not IsArray(Application(CookieName&"_blog_Category")) Then
	Dim log_CategoryList
	Set log_CategoryList=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT cate_ID,cate_Name,cate_Order FROM blog_Category ORDER BY cate_Order ASC"
	log_CategoryList.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If log_CategoryList.EOF And log_CategoryList.BOF Then
		Redim Arr_Category(3,0)
	Else
		Arr_Category=log_CategoryList.GetRows
	End If
	log_CategoryList.Close
	Set log_CategoryList=Nothing
	Application.Lock
	Application(CookieName&"_blog_Category")=Arr_Category
	Application.UnLock
Else
	Arr_Category=Application(CookieName&"_blog_Category")
End IF

'写入论坛板块列表
Dim Arr_Forums
If Not IsArray(Application(CookieName&"_blog_Forums")) Then
	Dim log_ForumList
	Set log_ForumList=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT forum_ID,forum_Name,forum_Order,forum_ThreadNums FROM blog_Forums ORDER BY forum_Order ASC"
	log_ForumList.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If log_ForumList.EOF And log_ForumList.BOF Then
		Redim Arr_Forums(4,0)
	Else
		Arr_Forums=log_ForumList.GetRows
	End If
	log_ForumList.Close
	Set log_ForumList=Nothing
	Application.Lock
	Application(CookieName&"_blog_Forums")=Arr_Forums
	Application.UnLock
Else
	Arr_Forums=Application(CookieName&"_blog_Forums")
End If

'写入表情符号
Dim Arr_Smilies
IF Not IsArray(Application(CookieName&"_blog_Smilies")) Then
	Dim log_SmiliesList
	Set log_SmiliesList=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT sm_ID,sm_Image,sm_Text FROM blog_Smilies ORDER BY sm_ID ASC"
	log_SmiliesList.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If log_SmiliesList.EOF And log_SmiliesList.BOF Then
		Redim Arr_Smilies(3,0)
	Else
		Arr_Smilies=log_SmiliesList.GetRows
	End If
	log_SmiliesList.Close
	Set log_SmiliesList=Nothing
	Application.Lock
	Application(CookieName&"_blog_Smilies")=Arr_Smilies
	Application.UnLock
Else
	Arr_Smilies=Application(CookieName&"_blog_Smilies")
End IF

'写入关键字列表
Dim Arr_Keywords
IF Not IsArray(Application(CookieName&"_blog_Keywords")) Then
	Dim log_KeywordsList
	Set log_KeywordsList=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT key_ID,key_Text,key_URL,key_Image FROM blog_Keywords ORDER BY key_ID ASC"
	log_KeywordsList.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If log_KeywordsList.EOF And log_KeywordsList.BOF Then
		Redim Arr_Keywords(4,0)
	Else
		Arr_Keywords=log_KeywordsList.GetRows
	End If
	log_KeywordsList.Close
	Set log_KeywordsList=Nothing
	Application.Lock
	Application(CookieName&"_blog_Keywords")=Arr_Keywords
	Application.UnLock
Else
	Arr_Keywords=Application(CookieName&"_blog_Keywords")
End IF

'写入首页链接列表
Dim Arr_Bloglinks
IF Not IsArray(Application(CookieName&"_blog_Bloglinks")) Then
	Dim log_BloglinksList
	Set log_BloglinksList=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT link_Name,link_URL,link_Image FROM blog_Links WHERE link_IsMain=True ORDER BY link_Order ASC"
	log_BloglinksList.Open SQL,Conn,1,1
	SQLQueryNums=SQLQueryNums+1
	If log_BloglinksList.EOF And log_BloglinksList.BOF Then
		Redim Arr_Bloglinks(3,0)
	Else
		Arr_Bloglinks=log_BloglinksList.GetRows
	End If
	log_BloglinksList.Close
	Set log_BloglinksList=Nothing
	Application.Lock
	Application(CookieName&"_blog_Bloglinks")=Arr_Bloglinks
	Application.UnLock
Else
	Arr_Bloglinks=Application(CookieName&"_blog_Bloglinks")
End IF

'写入词语屏蔽列表
Dim Arr_WordFilter
If Not IsArray(Application(CookieName&"_blog_WordFilter")) Then
	Dim log_WordFilter
	Set log_WordFilter=Server.CreateObject("ADODB.RecordSet")
	SQL="SELECT wf_ID,wf_String,wf_Replace FROM blog_WordFilter ORDER BY wf_ID ASC"
	log_WordFilter.Open SQL,Conn,1,1
	If log_WordFilter.Eof And log_WordFilter.Bof Then
		Redim Arr_WordFilter(3,0)
	Else
		Arr_WordFilter=log_WordFilter.GetRows
	End If
	log_WordFilter.Close
	Set log_WordFilter = Nothing
	Application.Lock
	Application(CookieName&"_blog_WordFilter")=Arr_WordFilter
	Application.UnLock
Else
	Arr_WordFilter=Application(CookieName&"_blog_WordFilter")
End If
%>
<!--#include file="ActiveUser.asp" -->