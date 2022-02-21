<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Dim log_Year,log_Month,log_Day,cateID,viewType,viewMode,sortBy,SQLFiltrate,Url_Add
	viewMode=Session("viewMode")'显示模式函数开始
    If Request.QueryString("viewMode")="list" Then
        viewMode="list"
        Session("viewMode")="list"
    ElseIf Request.QueryString("viewMode")="normal" Then
        viewMode="normal"
        Session("viewMode")=""
    End If
    viewType=CheckStr(Trim(Request.QueryString("viewType")))'显示模式函数结束
	log_Year=CheckStr(Trim(Request.QueryString("log_Year")))
	log_Month=CheckStr(Trim(Request.QueryString("log_Month")))
	log_Day=CheckStr(Trim(Request.QueryString("log_Day")))
	cateID=CheckStr(Trim(Request.QueryString("cateID")))
	
	SQLFiltrate="WHERE "
	Url_Add="?"
	
	sortBy=Session("sortBy")'显示模式函数开始
        If CheckStr(Trim(Request.QueryString("sortBy")))="" Then
                sortBy="log_IsTop DESC,log_ID"
                Session("sortBy")="log_IsTop ASC,log_ID"
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
        End If'显示模式函数结束
	
	IF IsInteger(cateID)=True Then
		SQLFiltrate=SQLFiltrate&" log_CateID="&CateID&" AND"
		Url_Add=Url_Add&"CateID="&CateID&"&"
	End IF

	IF IsInteger(log_Year)=True Then
		SQLFiltrate=SQLFiltrate&" log_PostYear="&log_Year&" AND"
		Url_Add=Url_Add&"log_Year="&log_Year&"&"
	End IF
	IF IsInteger(log_Month)=True Then
		SQLFiltrate=SQLFiltrate&" log_PostMonth="&log_Month&" AND"
		Url_Add=Url_Add&"log_Month="&log_Month&"&"
	End IF
	IF IsInteger(log_Day)=True Then
		SQLFiltrate=SQLFiltrate&" log_PostDay="&log_Day&" AND"
		Url_Add=Url_Add&"log_Day="&log_Day&"&"
	End IF
	Call MemberCenter
	Response.Write("<br>")
	Call Calendar(log_Year,log_Month,log_Day)
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call hotBlogList
	Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call blogSearch%>
        <div class="siderbar_head"><img src="images/sider_links.gif" border="0" align="absmiddle" /> 友情链接</div>
      <div class="siderbar_main">
        <%
		Dim blog_LinksListNums,blog_LinksListNumI,blog_LinksMainIMG,blog_LinksMainTXT
		blog_LinksListNums=Ubound(Arr_Bloglinks,2)
		For blog_LinksListNumI=0 To blog_LinksListNums
			IF Arr_Bloglinks(2,blog_LinksListNumI)<>Empty Then
				blog_LinksMainIMG=blog_LinksMainIMG&"<div class=""hyperlink""><a href="""&Arr_Bloglinks(1,blog_LinksListNumI)&""" target=""_blank""><img src="""&Arr_Bloglinks(2,blog_LinksListNumI)&""" border=""0"" alt="""&Arr_Bloglinks(0,blog_LinksListNumI)&"""></a></div>"
			Else
				blog_LinksMainTXT=blog_LinksMainTXT&"<div class=""hyperlink""><a href="""&Arr_Bloglinks(1,blog_LinksListNumI)&""" target=""_blank"">"&Arr_Bloglinks(0,blog_LinksListNumI)&"</a></div>"
			End IF
		Next
		Response.Write(blog_LinksMainIMG)
		Response.Write(blog_LinksMainTXT)
		%>
        <div align="right" style="margin: 4px;"><a href="bloglinks.asp">更多友情链接...</a></div>
      </div>
      <br>
        <div class="siderbar_head"><img src="images/sider_other.gif" border="0" align="absmiddle" /> 其他信息</div>
      <div class="siderbar_main" align="center"><a href="http://www.51pcbook.com/default.asp"></a><br />
            <img src="images/utf8.gif" border="0" alt="BLOG编码" align="absmiddle"><br />
            <%If IsInteger(cateID)=True Then
		  Response.Write("<a href=""blogrss1.asp?cateID="&cateID&""" target=""_blank"">")
		  Else
		  Response.Write("<a href=""blogrss1.asp"" target=""_blank"">")
		  End If
		  %>
        <img src="images/rss1.gif" border="0" alt="RSS 1.0" align="absmiddle"><br />
            <%If IsInteger(cateID)=True Then
		  Response.Write("<a href=""blogrss2.asp?cateID="&cateID&""" target=""_blank"">")
		  Else
		  Response.Write("<a href=""blogrss2.asp"" target=""_blank"">")
		  End If
		  %>
        <img src="images/rss2.gif" border="0" alt="RSS 2.0" align="absmiddle"><br />
            <a href="http://www.creativecommons.cn/licenses/by-nc-sa/1.0/" target="_blank"><img src="images/cc.gif" border="0" alt="创作共用协议" align="absmiddle"></a><br>
      </div>
      <br>
      <br />
    </td>
    <td width="100%" valign="top" background="topBar_bg.gif" bgcolor="#FFFFFF">
<div>
      <iframe border="0" frameborder="0" framespacing="0" height="16" marginheight="0" marginwidth="0" noresize scrolling="no" width="100%" vspale="0" src="blognews.asp"></iframe>
    </div>
        
      <%
Dim CurPage
If CheckStr(Request.QueryString("Page"))<>Empty Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
Else
	Curpage=1
End If

Dim webLog
Set webLog=Server.CreateObject("Adodb.Recordset")
SQL="SELECT L.*,C.cate_Name FROM blog_Content AS L,blog_Category AS C "&SQLFiltrate&" C.cate_ID=L.log_CateID ORDER BY log_IsTop DESC,log_ID DESC"
If viewMode="list" Then'显示模式修改开始
SQL="SELECT L.*,C.cate_Name FROM blog_Content AS L,blog_Category AS C "&SQLFiltrate&" C.cate_ID=L.log_CateID ORDER BY "&sortBy&" DESC"
End IF'显示模式修改结束
webLog.Open SQL,CONN,1,1
SQLQueryNums=SQLQueryNums+1
If webLog.EOF AND webLog.BOF Then 
	Response.Write("<div class=""message"">暂时没有日志</div>")
Else
	Dim log_Author,weblog_ID,log_IsShow,log_ShowURL,log_IsTop,log_Intro,log_Weather,log_tags
	If viewMode="list" Then blogPerPage=blogPerPage*6'显示模式修改：1表示倍数
	webLog.PageSize=blogPerPage
	webLog.AbsolutePage=CurPage
	Log_Num=webLog.RecordCount
	Dim Log_Num,MultiPages,PageCount
	MultiPages="<span class=""smalltxt"">"&MultiPage(Log_Num,blogPerPage,CurPage,Url_Add)&"</span>"
	Response.Write("<table width=""100%"" height=""5"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"&MultiPages&"</td><td align=right><font color=#FF0000>显示模式: </font><a href='"&URL_Add&"viewMode=normal'>正常</a> | <a href='"&URL_Add&"viewMode=list'>文章列表</a></td></tr></table>")
	Response.Write("<table width=""100%"" height=""9"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
	If viewMode="list" Then Response.Write("<TABLE cellSpacing=1 cellPadding=6 width='100%' align=center border=0 style='font-family: Verdana;font-size: 11px;color: #666666;'><tr><td colspan='4' align=center><div style='border-bottom: 1px dashed #CCCCCC;'><B>文章列表</B>　　　　　　　　　　　　　　　　　　排序方式:　<a href='"&URL_Add&"sortBy=log_PostTime'> 发表时间 </a>,<a href='"&URL_Add&"sortBy=log_CommNums'> 评论数 </a>,<a href='"&URL_Add&"sortBy=log_CateID'> 分类 </a>,<a href='"&URL_Add&"sortBy=log_ViewNums'> 查看数</a></div></td></tr>")
	Do Until webLog.EOF OR PageCount=blogPerPage
		weblog_ID=weblog("log_ID")
		log_IsShow=weblog("log_IsShow")
		log_Author=webLog("log_Author")
		log_IsTop=weblog("log_IsTop")
		log_Weather=Split(weblog("log_Weather"),"|")
		log_Intro=Replace(webLog("log_Intro"),"&#39;&#39;","&#39;")
		If viewMode="list" Then'显示模式修改开始
		  If log_IsShow = True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
			If log_IsTop=True Then
			Response.Write("<tr><td><img src=""images/blogtop.gif"" border=""0"" align=""absmiddle"" /><font color=""#FF0000""><strong>[置顶]</strong></font></span>  <A href='blogview.asp?logID="&webLog("log_ID")&"'>"&HTMLEncode(cutStr(webLog("log_Title"),34))&"</a></td><td><a href=""member.asp?action=view&memName="&log_Author&"""><B>"&log_Author&"</B></a></td><td>"&DateToStr(webLog("log_PostTime"),"Y-m-d H:I")&"</td><td><a href='blogview.asp?logID="&webLog("log_ID")&"#comment' title='评论'>"&webLog("log_CommNums")&"</a>|<a href=""trackback.asp?logID="&weblog_ID&""" title='引用'>"&webLog("log_QuoteNums")&"</a>|"&webLog("log_ViewNums")&"</strong></td></tr>")
			Else	
			Response.Write("<tr><td>[<A href='default.asp?CateID="&webLog("log_CateID")&"'><font color=#FF0000>"&webLog("cate_Name")&"</font></a>] - <A href='blogview.asp?logID="&webLog("log_ID")&"'>"&HTMLEncode(cutStr(webLog("log_Title"),34))&"</a></td><td><a href=""member.asp?action=view&memName="&log_Author&"""><B>"&log_Author&"</B></a></td><td>"&DateToStr(webLog("log_PostTime"),"Y-m-d H:I")&"</td><td><a href='blogview.asp?logID="&webLog("log_ID")&"#comment' title='评论'>"&webLog("log_CommNums")&"</a>|<a href=""trackback.asp?logID="&weblog_ID&""" title='引用'>"&webLog("log_QuoteNums")&"</a>|"&webLog("log_ViewNums")&"</td></tr>")
		  End IF
		  Else
			Response.Write("<tr><td><img src=""images/blog_ishidden.gif"" border=""0"" align=""absmiddle"" />&nbsp;&nbsp;<strong>这是一篇隐藏日志，只有管理员才能观看</strong></td><td><B>隐藏</B></a></td><td>****-**-** **:**</td><td>隐藏</td></tr>")
		  End If
		Else
			Response.Write("<div class=""content_head"">")
		If IsInteger(cateID)=False Then
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&""">"
		Else
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&"&cateID="&cateID&""">"
		End If
		If log_IsShow = True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
			If log_IsTop=True Then
				Response.Write("<span style=""cursor:hand;"" onclick=""showIntro('log_"&weblog_ID&"');"" title=""点击查看详细介绍""><img src=""images/blogtop.gif"" border=""0"" align=""absmiddle"" /><font color=""#FF0000""><strong>[置顶]</strong></font></span>  "&log_ShowURL&"<strong>"&HTMLEncode(cutStr(webLog("log_Title"),39))&"</strong></a>&nbsp;&nbsp;&nbsp;[ "&DateToStr(webLog("log_PostTime"),"Y-m-d")&" &nbsp;|&nbsp; <a href="""&HTMLEncode(webLog("log_FromURL"))&""" target=""_blank"">"&HTMLEncode(webLog("log_From"))&"</a> ]</div><div style=""display:none;"" id=""log_"&weblog_ID&""">")
			Else
				Response.Write("<img src=""images/weather/"&log_Weather(0)&".gif"" alt="""&log_Weather(1)&""" align=""absmiddle""> "&log_ShowURL&"<strong>"&HTMLEncode(cutStr(webLog("log_Title"),39))&"</strong></a>&nbsp;&nbsp;&nbsp;[ "&DateToStr(webLog("log_PostTime"),"Y-m-d")&" &nbsp;|&nbsp; <a href="""&webLog("log_FromURL")&""" target=""_blank"">"&webLog("log_From")&"</a> ]</div>")
			End IF
			if instr(log_Intro,"[mem]")>instr(log_Intro,"[/mem]") then
    if memName<>Empty then
    log_Intro=replace(replace(log_Intro,"[mem]",""),"[/mem]","")
    else            
    log_Intro=log_ShowURL&"<BR><p align=center>此日志包含会员才能看到的内容，请进入全文观看</a></p>"
    end if
    end if
			Response.Write("<div class=""content_main"">"&Ubbcode(log_Intro,weblog("log_DisSM"),weblog("log_DisUBB"),weblog("log_DisIMG"),weblog("log_AutoURL"),weblog("log_AutoKEY")))
			Response.Write("<br>")
			'显示TAGS
				 if ShowTag(weblog_ID,"Show")<>empty then
				    Response.Write (ShowTag(weblog_ID,"Show"))
				 end if
				'显示TAGS结束
				Response.Write("<br><br>")
			If HtmlEncode(webLog("log_Content"))<>log_Intro Then Response.Write(log_ShowURL&"<img src=""images/icon_readmore.gif"" align=""absmiddle"" border=""0""> 阅读全文……</a>")
			Response.Write("<div align=""right"" class=""smalltxt"" height=""32px"">作者：<a href=""member.asp?action=view&memName="&Server.URLEncode(log_Author)&""">"&log_Author&"</a>&nbsp;|&nbsp;分类：<a href=""default.asp?cateID="&webLog("log_CateID")&""">"&webLog("cate_Name")&"</a>&nbsp;|&nbsp;<a href=""blogview.asp?logID="&weblog_ID&"#comment"">评论："&webLog("log_CommNums")&"</a>&nbsp;|&nbsp;<a href=""trackback.asp?logID="&weblog_ID&""">引用："&webLog("log_QuoteNums")&"</a>&nbsp;|&nbsp;查看："&webLog("log_ViewNums")&"")
			If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then Response.Write("&nbsp;<a href=""blogedit.asp?logID="&weblog_ID&"""><img src=""images/icon_edit.gIf"" border=""0"" align=""absmiddle"" alt=""编辑日志""></a>")
			Response.Write("</div></div>")
			If log_IsTop=True Then Response.Write("</div>")
		Else
			Response.Write("<img src=""images/blog_ishidden.gif"" border=""0"" align=""absmiddle"" />&nbsp;&nbsp;<strong>这是一篇隐藏日志，只有管理员才能观看，请先登录</strong></div>")
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
	Response.Write("</table></div>")
        Response.Write(MultiPages)
Else
        Response.Write(MultiPages)
        Response.Write("</td></tr></table>")
End if
%>
    </td>
  </tr>
</table>
<!--#include file="footer.asp" -->