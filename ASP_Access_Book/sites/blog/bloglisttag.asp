<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/library.asp" -->
<%
dim server_v1,server_v2
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
   response.write "<center><div style=border: 2px solid #CCC;padding:20px;margin=50px;width:450px;>你提交的路径有误，禁止从站点外部提交数据请不要乱该参数！</div></center>"
   response.end
end if
%>
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#F7F7F4" nowrap>
<%Dim log_Year,log_Month,log_Day,cateID,SQLFiltrate,Url_Add,viewMode,sortBy,tag
	viewMode=Session("viewMode")
 	If CheckStr(Trim(Request.QueryString("viewMode")))="list" Then
  		viewMode="list"
  		Session("viewMode")="list"
 	ElseIf CheckStr(Trim(Request.QueryString("viewMode")))="normal" Then
  		viewMode="normal"
  		Session("viewMode")=""
 	End If
log_Year=CheckStr(Trim(Request.QueryString("log_Year")))
log_Month=CheckStr(Trim(Request.QueryString("log_Month")))
log_Day=CheckStr(Trim(Request.QueryString("log_Day")))
cateID=CheckStr(Trim(Request.QueryString("cateID")))
tag=CheckStr(Trim(Request.QueryString("tags")))
If IsInteger(cateID) = False Then cateID = 0
SQLFiltrate="WHERE "
Url_Add="?tags="&tag&"&"
If cateID<>0 Then
	SQLFiltrate=SQLFiltrate&" log_CateID="&CateID&" AND"
	Url_Add=Url_Add&"CateID="&CateID&"&"
End If
If IsInteger(log_Year)=True Then
	SQLFiltrate=SQLFiltrate&" log_PostYear="&log_Year&" AND"
	Url_Add=Url_Add&"log_Year="&log_Year&"&"
End If
If IsInteger(log_Month)=True Then
	SQLFiltrate=SQLFiltrate&" log_PostMonth="&log_Month&" AND"
	Url_Add=Url_Add&"log_Month="&log_Month&"&"
End If
If IsInteger(log_Day)=True Then
	SQLFiltrate=SQLFiltrate&" log_PostDay="&log_Day&" AND"
	Url_Add=Url_Add&"log_Day="&log_Day&"&"
End If
    sortBy=Session("sortBy")
	If CheckStr(Trim(Request.QueryString("sortBy")))="" Then
		sortBy="log_ID"'log_PostTime
		Session("sortBy")="log_ID"'log_PostTime
	ElseIf CheckStr(Trim(Request.QueryString("sortBy")))="log_ID" Then
		sortBy="log_ID"
		Session("sortBy")="log_ID"
	ElseIf CheckStr(Trim(Request.QueryString("sortBy")))="log_PostTime" Then
		sortBy="log_PostTime"
		Session("sortBy")="log_PostTime"
	ElseIf CheckStr(Trim(Request.QueryString("sortBy")))="log_CateID" Then
		sortBy="log_CateID"
		Session("sortBy")="log_CateID"
	ElseIf CheckStr(Trim(Request.QueryString("sortBy")))="log_ViewNums" Then
		sortBy="log_ViewNums"
		Session("sortBy")="log_ViewNums"
	ElseIf CheckStr(Trim(Request.QueryString("sortBy")))="log_CommNums" Then
		sortBy="log_CommNums"
		Session("sortBy")="log_CommNums"
	End If
	IF tag<>empty then
	   SQLFiltrate=SQLFiltrate&"tagsName='"&tag&"' AND"
	End IF%>
	<%Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call NewBlogList
    Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call blogSearch%><br>
	</td>
	<td width="100%" valign="top" bgcolor="#FFFFFF">
          <%IF tag=Empty Then
Response.Write("<div class=""message"">对不起，您没有提供Tag，无法为您提供数据。<br><br><a href=default.asp>请点击返回</a></div>")
Else
Dim CurPage
If CheckStr(Request.QueryString("Page"))<>Empty Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
Else
	Curpage=1
End If
Dim webLog
Set webLog=Server.CreateObject("Adodb.Recordset")
	If viewMode="list" Then
		SQL="SELECT L.*,C.cate_Name,A.* FROM blog_Content AS L,blog_Category AS C,blog_tag AS A "&SQLFiltrate&" C.cate_ID=L.log_CateID AND L.log_ID=A.blog_ID ORDER BY log_IsTop ASC,"&sortBy&" DESC"
	else
		SQL="SELECT L.*,C.cate_Name,A.* FROM blog_Content AS L,blog_Category AS C,blog_tag AS A "&SQLFiltrate&" C.cate_ID=L.log_CateID AND L.log_ID=A.blog_ID ORDER BY log_IsTop ASC,log_ID DESC"'log_PostTime DESC
	End IF
webLog.Open SQL,CONN,1,1
SQLQueryNums=SQLQueryNums+1
If webLog.EOF AND webLog.BOF Then 
	Response.Write("<div class=""message"">暂时没有日志</div>")
Else
	Dim log_Author,weblog_ID,log_IsShow,log_ShowURL,log_IsTop,log_Intro,log_Weather,log_mf,log_tags
	If viewMode="list" Then
	blogPerPage=blogPerPage*5
	End if
	webLog.PageSize=blogPerPage
	webLog.AbsolutePage=CurPage
	Log_Num=webLog.RecordCount
	Dim Log_Num,MultiPages,PageCount
	MultiPages="<span class=""smalltxt"">"&MultiPage(Log_Num,blogPerPage,CurPage,Url_Add)&"</span>"
	Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"&MultiPages&" "&Log_Num&" 篇日志使用了该 TAG <B><font color=#ff0000>"&CheckStr(Trim(Request.QueryString("tags")))&"</font></B></td><td align=""right""><font color=red>显示模式:</font> <a href="""&URL_Add&"viewMode=normal"">普通模式</a> | <a href="""&URL_Add&"viewMode=list"">列表模式</a></td></tr></table>")
	Response.Write("<table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
	If viewMode="list" Then
	Response.Write("<div class=""content_heads""><TABLE cellSpacing=1 cellPadding=5 width='588' align=center border=0 style='font-family: Verdana;font-size: 11px;color: #666666;'><tr><td colspan='4' ><div style='border-bottom: 1px dashed #cccccc;'><B>树形Blog列表</B>　　　　　　　　　　　　　　　　　　排序方式：<span class=""arrow"">8</span><a href="""&URL_Add&"sortBy=log_PostTime&tags="&tag&""">发表时间</a><span class=""arrow"">4</span><a href="""&URL_Add&"sortBy=log_ID&tags="&tag&""">分类</a><span class=""arrow"">4</span><a href="""&URL_Add&"sortBy=log_ViewNums&tags="&tag&""">查看数</a><span class=""arrow"">4</span><a href="""&URL_Add&"sortBy=log_CommNums&tags="&tag&""">评论数</a></div></td></tr>")'显示模式修改结束
	End IF
	Do Until webLog.EOF OR PageCount=blogPerPage
		weblog_ID=weblog("log_ID")
		log_IsShow=weblog("log_IsShow")
		log_Author=webLog("log_Author")
		log_IsTop=weblog("log_IsTop")
		log_Weather=Split(weblog("log_Weather"),"|")
		log_Intro=Replace(Replace(webLog("log_Intro"),"&#39;&#39;","&#39;"),"&nbsp;"," ")
		If viewMode="list" Then
		If log_IsShow=True Or (log_IsShow=False And (memStatus="SupAdmin" Or (memStatus="Admin" And memName=log_Author))) Then
		    Response.Write("<tr><td>[<A href=""default.asp?CateID="&webLog("log_CateID")&""">"&webLog("cate_Name")&"</a>] - <A href='blogview.asp?logID="&webLog("log_ID")&"' title="""&log_Author&" 发表于 "&DateToStr(webLog("log_PostTime"),"Y-m-d H:I A")&"&#13;&#10;&#13;&#10;"&webLog("log_Title")&""""">"&SplitLines(cutStr(Replace(webLog("log_Title"),"<br>",""),36),0)&"</a></td><td width=52><a href=""member.asp?action=view&memName="&log_Author&"""><B>"&log_Author&"</B></a></td><td width=124>"&DateToStr(webLog("log_PostTime"),"Y-m-d H:I A")&"</td><td><a href='blogview.asp?logID="&webLog("log_ID")&"#comment' title='评论'>"&webLog("log_CommNums")&"</a>|<a href=""trackback.asp?logID="&weblog_ID&""" title='引用'>"&webLog("log_QuoteNums")&"</a>|"&webLog("log_ViewNums")&"</td></tr>")
		End If
	Else
		Response.Write("<div class=""content_head"">")
		If IsInteger(cateID)=False Then
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&""">"
		Else
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&"&cateID="&cateID&""">"
		End If
		If log_IsTop=True Then
			Response.Write("<span style=""cursor:hand;"" onclick=""showIntro('log_"&weblog_ID&"');"" title=""点击查看详细介绍""><img src=""images/blogtop.gif"" border=""0"" align=""absmiddle"" /><font color=""#FF0000""><strong>[置顶]</strong></font></span>  "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;&nbsp;<font color=#615134>[ "&DateToStr(webLog("log_PostTime"),"Y-m-d")&" &nbsp;|&nbsp; </font><a href="""&webLog("log_FromURL")&""" target=""_blank"">"&webLog("log_From")&"</a> <font color=#615134>]</font></div><div style=""display:none;"" id=""log_"&weblog_ID&"""><div class=""content_main"">")
		Else
			Response.Write("<img src=""images/weather/"&log_Weather(0)&".gif"" alt="""&log_Weather(1)&""" align=""absmiddle""> "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;&nbsp;<font color=#615134>[ "&DateToStr(webLog("log_PostTime"),"Y-m-d")&" &nbsp;|&nbsp;</font> <a href="""&webLog("log_FromURL")&""" target=""_blank"">"&webLog("log_From")&"</a> <font color=#615134>]</font></div><div class=""content_main"">")
		End IF
		If log_IsShow = True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
			 if not isnull(log_mf) and log_mf<>"" then
           
end if

			Response.Write(Ubbcode(log_Intro,weblog("log_DisSM"),weblog("log_DisUBB"),weblog("log_DisIMG"),weblog("log_AutoURL"),weblog("log_AutoKEY")))
			Response.Write("<br><br>")
			'显示TAGS
				 if ShowTag(weblog_ID,"Show")<>empty then
				    Response.Write (ShowTag(weblog_ID,"Show"))
				 end if
				'显示TAGS结束
				Response.Write("<br><br>")
			If HtmlEncode(webLog("log_Content"))<>log_Intro Then Response.Write(log_ShowURL&"<img src=""images/icon_readmore.gif"" align=""absmiddle"" border=""0""> 阅读全文……</a>")
			Response.Write("<div align=""right"" class=""smalltxt"" height=""32px""><a href=""member.asp?action=view&memName="&log_Author&""">作者："&log_Author&"</a>&nbsp;|&nbsp;<a href=""default.asp?cateID="&webLog("log_CateID")&""">分类："&webLog("cate_Name")&"</a>&nbsp;|&nbsp;<a href=""blogview.asp?logID="&weblog_ID&"#comment"">评论："&webLog("log_CommNums")&"</a>&nbsp;|&nbsp;<a href=""trackback.asp?logID="&weblog_ID&""">引用："&webLog("log_QuoteNums")&"</a>&nbsp;|&nbsp;查看："&webLog("log_ViewNums")&"")
			If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then Response.Write("&nbsp;<a href=""blogedit.asp?logID="&weblog_ID&"""><img src=""images/icon_edit.gIf"" border=""0"" align=""absmiddle"" alt=""编辑日志""></a>")
			Response.Write("</div>")
		Else
			Response.Write("<b>对不起，本日志为隐藏日志，只有管理员或者作者可以查看！</b>")
		End If
		End If
		webLog.MoveNext
		PageCount=PageCount+1
	If viewMode="list" Then

Else
		Response.Write("</div></div>")
		If log_IsTop=True Then Response.Write("</div>")
		Response.Write("<table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
End if

	Loop
End If
	webLog.Close
Set webLog=Nothing
If viewMode="list" Then
	Response.Write("</table></div><br>")
        Response.Write(MultiPages)
Else
        Response.Write(MultiPages)
        Response.Write("</td></tr></table>")
End if
end if
%></td>
  </tr>
</table>
<!--#include file="footer.asp" -->