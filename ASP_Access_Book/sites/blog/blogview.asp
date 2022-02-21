<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<%Dim log_Year,log_Month,log_Day,logID,cateID,Url_Add
	 log_Year=CheckStr(Trim(Request.QueryString("log_Year")))
	 log_Month=CheckStr(Trim(Request.QueryString("log_Month")))
 	 log_Day=CheckStr(Trim(Request.QueryString("log_Day")))
 	 logID=CheckStr(Trim(Request.QueryString("logID")))
 	 cateID=CheckStr(Trim(Request.QueryString("cateID")))
	 Dim errMSG
	 If IsInteger(logID)=False OR (cateID<>Empty AND IsInteger(cateID)=False) Then
			errMSG="<br><center><div class=""msg_content""><a href=""default.asp"">对不起，无效的参数，点击返回首页重新操作</a></div></center>"
	 Else
		Dim log_View
		Set log_View=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT * FROM blog_Content WHERE log_ID="&logID&""
		log_View.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		If log_View.EOF AND log_View.BOF Then
	  		errMSG="<br><center><div class=""msg_content""><a href=""default.asp"">对不起，没有找到相关日志，点击返回首页重新操作</a></div></center>"
			log_View.Close
	  		Set log_View=Nothing
		Else
	  		Dim log_Title,log_Content,log_Author,log_PostTime,log_DisSM,log_DisUBB,log_DisIMG,log_AutoURL,log_From,log_FromURL,log_Modify,log_IsShow,log_QuoteNums,log_AutoKEY,log_DisComment,log_Weather,log_tags
			log_Title=log_View("log_Title")
			siteTitle = log_Title&" - "
			log_Content=log_View("log_Content")
			log_Author=log_View("log_Author")
			log_PostTime=log_View("log_PostTime")
			log_DisSM=log_View("log_DisSM")
			log_DisUBB=log_View("log_DisUBB")
			log_DisIMG=log_View("log_DisIMG")
			log_AutoURL=log_View("log_AutoURL")
			log_AutoKEY=log_View("log_AutoKEY")
			log_From=log_View("log_From")
			log_FromURL=log_View("log_FromURL")
			log_Modify=log_View("log_Modify")
			log_IsShow=log_View("log_IsShow")
			log_QuoteNums=log_View("log_QuoteNums")
			log_DisComment=log_View("log_DisComment")
			log_Weather=Split(log_View("log_Weather"),"|")
	  		log_tags=log_view("log_tags")
	  		log_View("log_ViewNums")=log_View("log_ViewNums")+1
			log_View.UPDATE
	  		log_View.Close
	  		Set log_View=Nothing
		End If
	End If%>
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call Calendar(log_Year,log_Month,log_Day)
	Response.Write("<br>")
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call blogSearch%>
      <br>


</td>
    <td width="100%" valign="top" bgcolor="#FFFFFF">
	
<%
If errMSG<>Empty Then
	Response.Write(errMSG)
Else
	Url_Add="?logID="&logID&"&"
	Dim cateQuery,catePage
	If cateID<>Empty Then
		cateQuery=" AND log_cateID="&cateID
		catePage="&cateID="&cateID
		Url_Add=Url_Add&"cateID="&cateID&"&"
	End If
	Dim log_PrevLog,log_PrevLogStr,log_PrevLogTitle,log_NextLog,log_NextLogStr,log_NextLogTitle
	Set log_PrevLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID,log_IsShow,log_Author FROM blog_Content WHERE log_ID<"&logID&""&cateQuery&" ORDER BY log_ID DESC")
	SQLQueryNums=SQLQueryNums+1
	If log_PrevLog.EOF AND log_PrevLog.BOF Then
		log_PrevLogStr=""
	Else
		If log_PrevLog(2) = True OR (log_PrevLog(2)=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_PrevLog(3)))) Then
			log_PrevLogTitle=log_PrevLog(0)
		Else
			log_PrevLogTitle="隐藏日志，无权浏览"
		End If
		log_PrevLogStr="<img src=""images/icon_ar.gif"" border=""0"" align=""absmiddle""> <a href='blogview.asp?logID="&log_PrevLog(1)&""&catePage&"'>"&log_PrevLogTitle&"</a>"
	End IF
	Set log_PrevLog=Nothing
	Set log_NextLog=Conn.Execute("SELECT TOP 1 log_Title,log_ID,log_IsShow,log_Author FROM blog_Content WHERE log_ID>"&logID&""&cateQuery&" ORDER BY log_ID ASC")
	SQLQueryNums=SQLQueryNums+1
	If log_NextLog.EOF AND log_NextLog.BOF Then
		log_NextLogStr=""
	Else
		If log_NextLog(2) = True OR (log_NextLog(2)=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_NextLog(3)))) Then
			log_NextLogTitle=log_NextLog(0)
		Else
			log_NextLogTitle="隐藏日志，无权浏览"
		End If
			log_NextLogStr="<a href=""blogview.asp?logID="&log_NextLog(1)&""&catePage&""">"&log_NextLogTitle&"</a> <img src=""images/icon_al.gif"" border=""0"" align=""absmiddle"">"
	End If
	Set log_PrevLog=Nothing%>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" style="margin-bottom:4px;">
        <tr>
          <td width="50%" align="left"><%=log_PrevLogStr%></td>
          <td width="50%" align="right"><%=log_NextLogStr%></td>
        </tr>
      </table>
      <%
		If log_IsShow = True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
			Response.Write("<div class=""content_head""><img src=""images/weather/"&log_Weather(0)&".gif"" border=""0"" align=""absmiddle"" alt="""&log_Weather(1)&"""> <b><font color=""#0066CC"">"&log_Title&"</font></b>&nbsp;&nbsp;&nbsp;[ 日期："&DateToStr(log_PostTime,"Y-m-d")&" ]&nbsp;&nbsp;&nbsp;[ 来自：<a href="""&HTMLEncode(log_FromURL)&""" target=""_blank"">"&HTMLEncode(log_From)&"</a> ]")
			If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then
				Response.Write("&nbsp;&nbsp;<a href=""blogedit.asp?logID="&logID&"""><img src=""images/icon_edit.gif"" border=""0"" align=""absmiddle"" alt=""编辑日志""></a>")
			End If%></div><div class="content_main"><%
			Response.Write("<SPAN id='fontzoom'>"&Ubbcodes(HTMLEncode(log_Content),log_DisSM,log_DisUBB,log_DisIMG,log_AutoURL,log_AutoKEY,logID)&"<br><br>")
			'显示TAGS
				 if ShowTag(logID,"Show")<>empty then
				    Response.Write (ShowTag(logID,"Show"))
				 end if				'显示TAGS结束
			IF log_Modify<>Empty Then
				Response.Write("<div align=""right"" class=""smalltxt"" height=""35px"">"&log_Modify&"</div>")
			End If
			Response.Write("<div align=""right"" class=""smalltxt"" height=""35px"">[ 阅读字体大小：<a href='javascript:fontZoom(16)' class=black><font color=red>大</font></a> <a href='javascript:fontZoom(14)' class=black><font color=red>中</font></a> <a href='javascript:fontZoom(12)' class=black><font color=red>小</font></a> ]</div>")
			if log_From="本站原创" then
			else
			response.write ("<div align=""left"" class=""smalltxt"" height=""35px""><b><font color=red>本站声明：</font></b><font color=Black>此文章来源于网络，如果未属名，可能因为此文被转摘多次，原作者不详，如果您认为侵权，请联系我。我将在第一时间按要求做出处理，并消除影响。</font></span></div>")
			end if
			Response.Write("<div><img src=""images/icon_trackback.gif"" border=""0"" align=""absmiddle""><b><a href=""trackback.asp?logID="&logID&""">引用通告地址</a> ("&log_QuoteNums&")：</b><br /><img src=""images/utf8.gif"" border=""0"" align=""absmiddle"" style=""cursor:hand;"" alt=""复制引用地址"" onClick=""CopyText(document.all.tbUTFURL)""><span class=""smalltxt"" id=""tbUTFURL"">"&siteURL&"/</span><br /><img src=""images/gbk.gif"" border=""0"" align=""absmiddle"" style=""cursor:hand;"" alt=""复制引用地址"" onClick=""CopyText(document.all.tbGBKURL)""><span class=""smalltxt"" id=""tbGBKURL"">"&siteURL&"</span></div></div>")
		  Else
				Response.Write("<br><center><div class=""msg_content""><a href=""default.asp"">对不起，你没有权限查看此日志，请点击返回首页</a></div></center>")
		  End If
		  %><table height="10"><tr><td><div>反向链接列表<script language="javascript" src="ref/feeler.asp?showType=1"></script></div></td></tr></table>
      <%
	  If log_IsShow=True OR (log_IsShow=False And (memStatus="SupAdmin" OR (memStatus="Admin" And memName=log_Author))) Then
		  Dim CurPage
		  If CheckStr(Request.QueryString("Page"))<>Empty Then
			  Curpage=CheckStr(Request.QueryString("Page"))
			  If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
		  Else
			  Curpage=1
		  End If
		  Dim blog_Comment
		  Set blog_Comment=Server.CreateObject("Adodb.RecordSet")
		   SQL="SELECT comm_ID,comm_Content,comm_Author,comm_PostTime,comm_DisSM,comm_DisUBB,comm_DisIMG,comm_AutoURL,comm_PostIP,comm_AutoKEY FROM blog_Comment WHERE blog_ID="&logID&" UNION ALL SELECT 0,tb_Intro,tb_Title,tb_PostTime,tb_URL,tb_Site,tb_ID,0,'127.0.0.1',0 FROM blog_Trackback WHERE blog_ID="&logID&" ORDER BY comm_PostTime DESC"
		  blog_Comment.Open SQL,Conn,1,1
		  SQLQueryNums=SQLQueryNums+1
		  If blog_Comment.EOF And blog_Comment.BOF Then
			Response.Write("<div class=""content_head"">暂时没有评论</div>")
		  Else
			Dim Comm_Nums,MultiPages,PageCount
			blog_Comment.PageSize=blogPerPage
			blog_Comment.AbsolutePage=CurPage
			Comm_Nums=blog_Comment.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(Comm_Nums,blogPerPage,CurPage,Url_Add)&"</span>"
			Response.Write(MultiPages)
			Response.Write("<table width=""100%"" height=""5"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
			Do Until blog_Comment.EOF OR PageCount=blogPerPage
				Dim blog_CommID,blog_CommAuthor,blog_CommContent
				blog_CommID=blog_Comment("comm_ID")
				blog_CommAuthor=blog_Comment("comm_Author")
				blog_CommContent=blog_Comment("comm_Content")
				Response.Write("<div class=""content_head""><a name=""commmark_"&blog_CommID&"""></a>")
				IF blog_CommID=0 Then
					Response.Write("<img src=""images/icon_ctb.gif"" border=""0"" align=""absmiddle""/>&nbsp;<span class=""smalltxt""><strong>引用通告：<a href="""&blog_Comment("comm_DisSM")&""" target=""_blank"">"&blog_Comment("comm_DisUBB")&"</a> 于 "&DateToStr(blog_Comment("comm_PostTime"),"Y-m-d H:I A")&"</strong></span>")
				Else
					Response.Write("<img src=""images/icon_quote.gif"" alt=""引用这个评论"" border=""0"" align=""absmiddle"" style=""cursor:hand;"" onclick=""blogquote('quote_"&blog_CommID&"','"&blog_CommAuthor&"','"&DateToStr(blog_Comment("comm_PostTime"),"Y-m-d H:I A")&"')"";/>")
					Response.Write("&nbsp;<span class=""smalltxt""><strong><a href=""member.asp?action=view&memName="&Server.URLEncode(blog_CommAuthor)&""">"&blog_CommAuthor&"</a> 于 "&DateToStr(blog_Comment("comm_PostTime"),"Y-m-d H:I A")&" 发表评论：</strong></span>")
				End IF
				If (memStatus="Admin" AND memName=log_Author) OR memStatus="SupAdmin" Then
					If blog_CommID=0 Then
						Response.Write("&nbsp;<a href=""trackback.asp?action=deltb&logID="&logID&"&tbID="&blog_Comment("comm_DisIMG")&""" title=""删除引用通告"" onClick=""winconfirm('你真的要删除这个引用吗？','trackback.asp?action=deltb&logID="&logID&"&tbID="&blog_Comment("comm_DisIMG")&"'); return false""><b><font color='#FF0000'>×</font></b></a>")
					Else
						Response.Write("&nbsp;<a href=""blogcomm.asp?action=delecomm&logID="&logID&"&commID="&blog_CommID&""" title=""删除评论"" onClick=""winconfirm('你真的要删除这个评论吗？','blogcomm.asp?action=delecomm&logID="&logID&"&commID="&blog_CommID&"'); return false""><b><font color=""#FF0000"">×</font></b></a>&nbsp;|&nbsp;<a href=""ipview.asp?ipdata="&blog_Comment("comm_PostIP")&""" title=""点击查看IP地址："&blog_Comment("comm_PostIP")&" 的来源"" target=""_blank"">IP</a>")
					End If
				End If
				If blog_CommID=0 Then
					Response.Write("&nbsp;</div><div class=""content_main""><b>标题：</b>"&blog_CommAuthor&"<br><b>链接：</b><a href="""&blog_Comment("comm_DisSM")&""" target=""_blank"">"&blog_Comment("comm_DisSM")&"</a><br><b>摘要：</b>"&blog_CommContent&"</div><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF"" border=""0""/><br>")
				Else
					Response.Write("&nbsp;</div><div class=""content_main""><div style=""display:none;"" id=""quote_"&blog_CommID&""">"&cutStr(DelQuote(Replace(HTMLEncode(blog_CommContent),"<br>","")),60)&"</div>"&UbbCode(HTMLEncode(blog_CommContent),blog_Comment("comm_DisSM"),blog_Comment("comm_DisUBB"),blog_Comment("comm_DisIMG"),blog_Comment("comm_AutoURL"),blog_Comment("comm_AutoKEY"))&"</div><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF"" border=""0""/><br>")
				End If
				PageCount=PageCount+1
				blog_Comment.MoveNext
			Loop
		  End If
		  blog_Comment.Close
		  Set blog_Comment=Nothing
		  Response.Write(MultiPages)
		  Response.Write("<table width=""100%"" height=""5"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
		  If log_DisComment=False Or (log_DisComment=True And (memStatus="SupAdmin" Or (memStatus="Admin" And memName=log_Author))) Then%>
	  <script language="JavaScript" src="include/ubbcode.js"></script><table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td colspan="2" bgcolor="#EFEFEF"><a name="#comment"></a><b>发表评论 - 不要忘了输入验证码哦！</b></td>
        </tr><form name="input" method="post" action="blogcomm.asp?action=postcomm">
        <tr>
          <td width="102" align="right" bgcolor="#FFFFFF" nowrap><b>作者：</b></td>
          <td width="100%" bgcolor="#FFFFFF">用户：
            <%If memName<>Empty Then
			Response.Write("<input name=""comm_memName"" type=""text"" id=""comm_memName"" value="""&memName&""" size=""12"" readonly />")
			Else
			Response.Write("<input name=""comm_memName"" type=""text"" id=""comm_memName"" size=""10"" />&nbsp;密码： <input name=""comm_memPassword"" type=""password"" id=""comm_memPassword"" size=""10"" />&nbsp;<input name=""comm_SaveMem"" type=""checkbox"" id=""comm_SaveMem"" value=""1"" /> 注册？")
			End IF%>&nbsp;验证：<input name="validatecode" type="text" id="validatecode" size="3" />&nbsp;<img src="include/validatecode.asp" align="absmiddle" border="0" /></td>
        </tr>
        <tr>
          <td align="right" valign="top" bgcolor="#FFFFFF"><b>评论：
          </b><div style="padding-left:5px;" align="left" width="100%">
                <br>
                <input name="comm_DisSM" type="checkbox" id="comm_DisSM" value="1" />
                禁止表情<br>
                <input name="comm_DisUBB" type="checkbox" id="comm_DisUBB" value="1" />
                禁止UBB<br>
                <input name="comm_DisIMG" type="checkbox" id="comm_DisIMG" value="1" />
                禁止图片<br>
                <input name="comm_AutoURL" type="checkbox" id="comm_AutoURL" value="1" checked />
                识别链接<br>
				<input name="comm_AutoKEY" type="checkbox" id="comm_AutoKEY" value="1" />
                识别关键字
            </div></td>
          <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td valign="top" nowrap><textarea name="message" cols="62" rows="8" wrap="VIRTUAL" id="Message" onSelect="javascript: storeCaret(this);" onClick="javascript: storeCaret(this);" onKeyUp="javascript: storeCaret(this);" onKeyDown="javascript: ctlent();"></textarea></td>
                <td width="122" align="left" valign="top"><div class="siderbar_head"><b>表&nbsp;&nbsp;情</b></div><div class="siderbar_main"><%Dim log_SmiliesListNums,log_SmiliesListNumI
			log_SmiliesListNums=Ubound(Arr_Smilies,2)
			TempVar=""
			For log_SmiliesListNumI=0 To log_SmiliesListNums
				Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" />")
				TempVar=" "
			Next
			%></div></td>
              </tr>
            </table><iframe border="0" frameBorder="0" frameSpacing="0" height="21" marginHeight="0" marginWidth="0" noResize scrolling="no" width="100%" vspale="0" src="attachment.asp"></iframe></td>
        </tr>
        <tr align="center">
          <td colspan="2" nowrap bgcolor="#FFFFFF"><input name="blog_ID" type="hidden" id="blog_ID" value="<%=logID%>" /><input type="submit" name="replysubmit" value=" 发表评论 [可按 Ctrl+Enter 发布] " onClick="this.disabled=true;document.input.submit();" />
&nbsp; <input name="Reset" type="reset" id="Reset" value=" 重置评论 " /></td>
          </tr></form>
      </table>
	  <%End If
	  End If
	  End If%>
</td>
  </tr>
</table>
<!--#include file="footer.asp" -->
<script>
function fontZoom(size)
{
document.getElementById('fontzoom').style.fontSize=size+'px'
}
</script>