<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" -->
<div class="GuoBlog_Main">
  <div class="GuoBlog_Left">
<%
	Dim log_Year,log_Month,log_Day,cateID,SQLFiltrate,Url_Add,mfyn,viewMode,url_view,tag
	log_Year=CheckStr(Trim(Request.QueryString("log_Year")))
	log_Month=CheckStr(Trim(Request.QueryString("log_Month")))
	log_Day=CheckStr(Trim(Request.QueryString("log_Day")))
	cateID=CheckStr(Trim(Request.QueryString("cateID")))
	tag=CheckStr(Trim(Request.QueryString("tags")))
	
	viewMode=session("viewMode")

	SQLFiltrate="WHERE "
	Url_Add="?"
	if viewMode="" or viewMode=empty then
	   viewMode="viewNormal"
	end if
	
	If Request.QueryString("viewMode")="viewList" Then
        viewMode="viewList"
        Session("viewMode")="viewList"
    ElseIf Request.QueryString("viewMode")="viewNormal" Then
        viewMode="viewNormal"
        Session("viewMode")="viewNormal"
    End If
	
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
	If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then
	ELSE
	   SQLFiltrate=SQLFiltrate&" log_mode<>'草稿' AND "
	End IF
	IF tag<>empty then
	   SQLFiltrate=SQLFiltrate & " tagsName = '"&tag&"' AND "
	End IF
    Call MemberCenter
	Response.Write("<br>")
	Call Calendar(log_Year,log_Month,log_Day)
	Response.Write("<br>")
	Call topblogs
	Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call blogSearch
	%>
      <br>
  </div>
	 <div class="GuoBlog_Right">
<div class="GuoBlog_SearchTag">
	     <form action="TagsSearch.asp" method="get" name="SearchTag">
		   <input name="tags" type="text" size="40" maxlength="50"><input type="submit" value=" Search Tags ">
		 </form>
  </div>
<%
IF tag=Empty Then
	Response.Write("<div class=""message"">对不起，您没有提供Tag，无法为您提供数据。<br><br><a href=default.asp>请点击返回</a></div>")
Else
Dim CurPage
If CheckStr(Request.QueryString("Page"))<>Empty and isnumeric(Request.QueryString("Page")) Then
	Curpage=CheckStr(Request.QueryString("Page"))
	If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
Else
	Curpage=1
End If

Dim webLog
Set webLog=Server.CreateObject("Adodb.Recordset")

'SQL="SELECT L.*,C.cate_Name FROM blog_Content AS L,blog_Category AS C "&SQLFiltrate&" C.cate_ID=L.log_CateID ORDER BY log_IsTop DESC,log_ID DESC"
SQL="SELECT blog_Content.*, blog_Category.cate_Name AS cate_Name,Blog_tag.TagsName AS TagsName FROM blog_Content LEFT OUTER JOIN blog_Category ON blog_Content.log_CateID = blog_Category.cate_ID LEFT OUTER JOIN Blog_tag ON blog_Content.log_ID = Blog_tag.Blog_ID "&SQLFiltrate&" 1=1 ORDER BY blog_Content.log_IsTop DESC, blog_Content.log_ID DESC"
webLog.Open SQL,CONN,1,1
SQLQueryNums=SQLQueryNums+1
If webLog.EOF AND webLog.BOF Then 
	Response.Write("<div class=""message"">暂时没有 <b>"&tag&"</b> 此标签的日志！</div>")
Else
	Dim log_Author,weblog_ID,log_IsShow,log_ShowURL,log_IsTop,log_Intro,log_Weather,log_modes,log_mf,log_key,log_tags
	log_key="B"
	if viewMode="viewList" then
	   blogPerPage=blogPerPage*8
	end if
	webLog.PageSize=blogPerPage
	webLog.AbsolutePage=CurPage
	Log_Num=webLog.RecordCount
    Response.Write ("<div class=""GuoBlog_tagsT"">")
    Response.Write ("<span class=""GuoBlog_Tags_font"">&nbsp;&nbsp;<a href=""tags.asp"">Tags</a>:"&tag&"</span>")
    Response.Write ("<br /><br />&nbsp;&nbsp;&nbsp;共有 "&Log_Num&" 篇日志使用了该Tag </div>")
	Dim Log_Num,MultiPages,PageCount,pageurls
	if log_Num>blogperpage then
	MultiPages="<table width=""98%""  border=""0"" cellspacing=""0"" cellpadding=""5""><tr><td width=""73%"">"&MultiPage(Log_Num,blogPerPage,CurPage,Url_Add&"tags="&tag&"&")&"</td><td width=""25%""><div align=""right""><img src=""images/normal.gif"" alt=""按摘要模式显示"" align=""absmiddle"" /> <a href="""&url_add&"viewMode=viewNormal&tags="&tag&""">摘要模式</a> | <img src=""images/list.gif"" alt=""按列表模式显示"" align=""absmiddle"" /> <a href="""&url_add&"viewMode=viewList&tags="&tag&""">列表模式</a></div></td></tr></table>"
	'MultiPages="<div class=""smalltxt"">"&MultiPage(Log_Num,blogPerPage,CurPage,Url_Add)&"</div> <a href="&url_view&">摘要模式</a> | <a href="&url_view&">列表模式</a>"
	Response.Write(MultiPages)
	else
	response.write "<table width=""98%""  border=""0"" cellspacing=""0"" cellpadding=""5""><tr><td width=""73%""> </td><td width=""25%""><div align=""right""><img src=""images/normal.gif"" alt=""按摘要模式显示"" align=""absmiddle"" /> <a href="""&url_add&"viewMode=viewNormal&tags="&tag&""">摘要模式</a> | <img src=""images/list.gif"" alt=""按列表模式显示"" align=""absmiddle"" /> <a href="""&url_add&"viewMode=viewList&tags="&tag&""">列表模式</a></div></td></tr></table>"
	end if
	if viewMode="viewList" then
	response.write "<table width=""98%""  border=""0"" cellspacing=""1"" cellpadding=""3"" bgcolor=""#cccccc""><tr><td width=""10%"" height=""35"" background=""images/blog_title_bg.gif"" bgcolor=""#ffffff""><div align=""center"" style=""font-weight: bold"">分类</div></td><td width=""56%"" background=""images/blog_title_bg.gif"" bgcolor=""#ffffff""><div align=""center"" style=""font-weight: bold"">标题</div></td><td width=""8%"" background=""images/blog_title_bg.gif"" bgcolor=""#ffffff""><div align=""center"" style=""font-weight: bold"">作者</div></td><td width=""15%"" background=""images/blog_title_bg.gif"" bgcolor=""#ffffff""><div align=""center"" style=""font-weight: bold"">发布时间</div></td><td width=""11%"" background=""images/blog_title_bg.gif"" bgcolor=""#ffffff""><div align=""center"" style=""font-weight: bold"">点击/回复</div></td></tr>"
    Do Until webLog.EOF OR PageCount=blogPerPage
	    weblog_ID=weblog("log_ID")
		log_IsShow=weblog("log_IsShow")
		log_Author=webLog("log_Author")
		log_IsTop=weblog("log_IsTop")
		log_modes=weblog("log_mode")
		log_mf=weblog("log_mf")
		log_key=weblog("log_key")
		log_tags=weblog("log_tags")
		'log_key="d"
		log_Weather=Split(weblog("log_Weather"),"|")
		log_Intro=Replace(webLog("log_Intro"),"&#39;&#39;","&#39;")

	   If IsInteger(cateID)=False Then
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&""">"
			pageurls="blogview.asp?logID="&weblog_ID
		Else
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&"&cateID="&cateID&""">"
			pageurls="blogview.asp?logID="&weblog_ID&"&cateID="&cateID
		End If
		
		'首页显示分页
		
		pcounts = ubound(split(weblog("log_content"),"[page][/page]"))
		If pcounts>0 then
		   paa="<img src=""images/page.gif"" alt=""PAGES"" align=""absmiddle"" />"
		   If pcounts>4 then
		   For pis=1 to 4
		       paa = paa & " <a href=""" & pageurls & "&pages=" & pis & """>" & pis &"</a>"
		   Next
		       paa = paa & " ... <a href=""" & pageurls & "&pages=" & pis & """>" & pcounts & "</a>"
		   Else
		   For pis=1 to pcounts
		       paa = paa & " <a href=""" & pageurls & "&pages=" & pis & """>" & pis &"</a>"
		   Next
		   End if
		End if
		'首页显示分页结束
		
		
       response.write "<tr><td bgcolor=""#ffffff""><div align=""center"">[<a href=""default.asp?cateID="&webLog("log_CateID")&""">"&webLog("cate_Name")&"</a>]</div></td><td bgcolor=""#ffffff"">"
		If not isnull(log_key) and (memStatus<>"SupAdmin" OR (memStatus<>"Admin" And memName<>log_Author)) then
		   Response.Write("<form name=""passblog"" method=""post"" action=""blogview.asp?logID="&weblog_ID&"""><b>对不起，查看此日志需要密码！</b>密码：<input name=""blogpass"" type=""password"" size=""6"" maxlength=""15""><input type=""submit"" name=""Submit"" value=""提交""></form>")
        else
		If log_IsTop=True Then
			Response.Write("<span style=""cursor:hand;"" onclick=""showIntro('log_"&weblog_ID&"');"" title=""点击查看详细介绍""><img src=""images/blogtop.gif"" border=""0"" align=""absmiddle"" /><font color=""#FF0000""><strong>[置顶]</strong></font></span>  "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;"&paa)
		Else
		'DateToStr(webLog("log_PostTime"),"Y-m-d")
			Response.Write("<img src=""images/weather/"&log_Weather(0)&".gif"" alt="""&log_Weather(1)&""" align=""absmiddle"" /> "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;"&paa)
		End IF
	   end if
	   Response.Write "</td><td bgcolor=""#ffffff""><a href=""member.asp?action=view&memName="&log_Author&""">"&log_Author&"</a></td><td bgcolor=""#ffffff"">"&DateToStr(webLog("log_PostTime"),"Y-m-d A")&"</td><td bgcolor=""#ffffff"">"&webLog("log_ViewNums")&" | "&webLog("log_CommNums")&"</td></tr>"
	   weblog.movenext
	   paa=""
	   log_tags=""
	   PageCount=PageCount+1
    loop
	response.write "</table>"
	else
	Do Until webLog.EOF OR PageCount=blogPerPage
		weblog_ID=weblog("log_ID")
		log_IsShow=weblog("log_IsShow")
		log_Author=webLog("log_Author")
		log_IsTop=weblog("log_IsTop")
		log_modes=weblog("log_mode")
		log_mf=weblog("log_mf")
		log_key=weblog("log_key")
		IF weblog("log_tags")=empty then
		log_tags=""
		ELSE
		log_tags=weblog("log_tags")
		End IF
		'log_key="d"
		log_Weather=Split(weblog("log_Weather"),"|")
		log_Intro=Replace(webLog("log_Intro"),"&#39;&#39;","&#39;")
		
		Response.Write("<div class=""content_heads""><div class=""content_head"">")
		If not isnull(log_key) and (memStatus<>"SupAdmin" OR (memStatus<>"Admin" And memName<>log_Author)) then
		   Response.Write("<form name=""passblog"" method=""post"" action=""blogview.asp?logID="&weblog_ID&"""><b>对不起，查看此日志需要密码，请输入正确的密码查看！</b><input name=""blogpass"" type=""password"" size=""15"" maxlength=""15""><input type=""submit"" name=""Submit"" value=""提交""></form>")
		else
		
		If log_IsShow = True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
		If IsInteger(cateID)=False Then
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&""">"
			pageurls="blogview.asp?logID="&weblog_ID
		Else
			log_ShowURL="<a href=""blogview.asp?logID="&weblog_ID&"&cateID="&cateID&""">"
			pageurls="blogview.asp?logID="&weblog_ID&"&cateID="&cateID
		End If
		
		'首页显示分页
		Dim pcounts,paa,pis
		pcounts = ubound(split(weblog("log_content"),"[page][/page]"))
		If pcounts>0 then
		   paa="<img src=""images/page.gif"" alt=""PAGES"" align=""absmiddle"" />"
		   If pcounts>4 then
		   For pis=1 to 4
		       paa = paa & " <a href=""" & pageurls & "&pages=" & pis & """>" & pis &"</a>"
		   Next
		       paa = paa & " ... <a href=""" & pageurls & "&pages=" & pis & """>" & pcounts & "</a>"
		   Else
		   For pis=1 to pcounts
		       paa = paa & " <a href=""" & pageurls & "&pages=" & pis & """>" & pis &"</a>"
		   Next
		   End if
		End if
		'首页显示分页结束
		
		If log_IsTop=True Then
			Response.Write("<span style=""cursor:hand;"" onclick=""showIntro('log_"&weblog_ID&"');"" title=""点击查看详细介绍""><img src=""images/blogtop.gif"" border=""0"" align=""absmiddle"" /><font color=""#FF0000""><strong>[置顶]</strong></font></span>  "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;"&paa&"&nbsp;&nbsp;&nbsp;[ "&DateToStr(webLog("log_PostTime"),"Y-m-d A")&" &nbsp;|&nbsp; <a href="""&webLog("log_FromURL")&""" target=""_blank"">"&webLog("log_From")&"</a> ]</div><div style=""display:none;"" id=""log_"&weblog_ID&"""><div class=""content_main"">")
		Else
		'DateToStr(webLog("log_PostTime"),"Y-m-d")
			Response.Write("<img src=""images/weather/"&log_Weather(0)&".gif"" alt="""&log_Weather(1)&""" align=""absmiddle"" /> "&log_ShowURL&"<strong>"&webLog("log_Title")&"</strong></a>&nbsp;&nbsp;"&paa&"&nbsp;&nbsp;&nbsp;[ "&DateToStr(webLog("log_PostTime"),"Y-m-d A")&" ]</div><div class=""content_main"">")
		End IF
		   if not isnull(log_mf) and log_mf<>"" then
		       Response.Write ("<DIV class=magicface><IMG title=点击观看魔法表情 style=""CURSOR: hand"" onclick=""ShowMagicFace('"&weblog("log_mf")&"');"" src=""magicface/images/"&weblog("log_mf")&".gif"" align=absMiddle><BR>魔法表情</DIV>")
               mfyn="yes" 
		   end if
			Response.Write(Ubbcode(log_Intro,weblog("log_DisSM"),weblog("log_DisUBB"),weblog("log_DisIMG"),weblog("log_AutoURL"),weblog("log_AutoKEY")))
			'Response.Write("<br/><br/>")
			    'If HtmlEncode(webLog("log_Content"))<>log_Intro Then Response.Write(log_ShowURL&"<img src=""images/icon_readmore.gif"" align=""absmiddle"" border=""0""> 阅读全文……</a>")
                '显示TAGS
				 if ShowTag(weblog_ID,"Show")<>empty then
				    Response.Write (ShowTag(weblog_ID,"Show"))
				 end if
				'显示TAGS结束
			    dim cg
				if log_modes="草稿" then
				   cg="<img src=images/sider_other.gif border=""0"" />此日志为草稿"
				else
				   cg=""
				end if
				Response.Write("<div align=""right"" class=""smalltxt"" height=""32px"">"&cg&"&nbsp;来源：<a href="""&webLog("log_FromURL")&""" target=""_blank"">"&webLog("log_From")&"</a> | 作者：<a href=""member.asp?action=view&memName="&log_Author&""">"&log_Author&"</a>&nbsp;|&nbsp;分类：<a href=""default.asp?cateID="&webLog("log_CateID")&""">"&webLog("cate_Name")&"</a>&nbsp;|&nbsp;<a href=""blogview.asp?logID="&weblog_ID&"#comment"">评论："&webLog("log_CommNums")&"</a>&nbsp;|&nbsp;<a href=""trackback.asp?logID="&weblog_ID&""">引用："&webLog("log_QuoteNums")&"</a>&nbsp;|&nbsp;查看："&webLog("log_ViewNums")&"")
		        If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then Response.Write("&nbsp;<a href=""blogedit.asp?logID="&weblog_ID&"""><img src=""images/icon_edit.gIf"" border=""0"" align=""absmiddle"" alt=""编辑日志""></a>")
			    Response.Write("</div>")
		Else
			   Response.Write("<b>对不起，本日志为隐藏日志，只有管理员或者作者可以查看！</b>")
		End If
		end if
		webLog.MoveNext
		PageCount=PageCount+1
		Response.Write("</div></div>")
		If log_IsTop=True Then Response.Write("</div>")
		if webLog.EOF OR PageCount=blogPerPage then
		   else
		   Response.Write("<br/>")
		end if
		paa=""
		log_tags=""
	Loop
	end if
End If
webLog.Close
Set webLog=Nothing
'判断是否有分页
If log_Num>blogperpage then
    Response.Write(MultiPages)
End If
End if
%>
	 </div>
</div>
<%
if mfyn="yes" then
%>
<DIV id=MagicFace style="Z-INDEX: 99; VISIBILITY: hidden; POSITION: absolute"></DIV>
<SCRIPT language=JavaScript>
			function MM_showHideLayers()
			{
			var i,p,v,obj,args=MM_showHideLayers.arguments;obj=document.getElementById("MagicFace");
			for (i=0; i<(args.length-2); i+=3) 
			if (obj) { v=args[i+2];
			if (obj.style) 
			{ obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }obj.visibility=v; }
			}
			function ShowMagicFace(MagicID)
			{var MagicFaceUrl = "magicface/flash/" + MagicID + ".swf";
			document.getElementById("MagicFace").innerHTML = '<object codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" width="500" height="350"><param name="movie" value="'+ MagicFaceUrl +'"><param name="menu" value="false"><param name="quality" value="high"><param name="play" value="false"><param name="wmode" value="transparent"><embed src="' + MagicFaceUrl +'" wmode="transparent" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="500" height="350"></embed></object>';
			document.getElementById("MagicFace").style.top = (document.body.scrollTop+((document.body.clientHeight-300)/2))+"px";
			document.getElementById("MagicFace").style.left = (document.body.scrollLeft+((document.body.clientWidth-480)/2))+"px";
			document.getElementById("MagicFace").style.visibility = 'visible';
			MagicID += Math.random();
			setTimeout("MM_showHideLayers('MagicFace','','hidden')",8000);
			var NowMeID = MagicID;
			}
</SCRIPT>
<%end if%>
<!--#include file="footer.asp" -->