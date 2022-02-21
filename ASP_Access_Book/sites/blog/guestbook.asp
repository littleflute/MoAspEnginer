<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Dim SQLFiltrate,Url_Add,gbAuthor
	gbAuthor=CheckStr(Request.QueryString("memName"))
	Url_Add="?"
	If gbAuthor<>Empty Then
		SQLFiltrate="WHERE gb_Author='"&gbAuthor&"'"
		Url_Add="?memName="&gbAuthor&""
	End If
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call NewCommList
	Response.Write("<br>")
	Call blogSearch%>
      <br>
      <br>
      <br>
</td>
    <td width="100%" valign="top" bgcolor="#FFFFFF">
	
<%
Dim msg_Title,msg_Content
If Request.QueryString("action")="postgb" Then
	If DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<15 Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">你发表留言速度太快了，点击返回上一页</a>"
	ElseIf Trim(Request.Form("validatecode"))=Empty Or Trim(Session("L-Blog_ValidateCode"))<>Trim(Request.Form("validatecode")) Then
        msg_Title="出现错误"
        msg_Content="<a href=""javascript:history.go(-1);"">请输入发表评论按钮旁边的验证码框，点击返回上一页</a>"
	Else
		Dim gb_AllreadyMem,gb_AllreadyMemErr
		Set gb_AllreadyMem=Server.CreateObject("ADODB.RecordSet")
		SQL="SELECT mem_Name,mem_Password,mem_Status,mem_LastIP FROM blog_Member WHERE mem_Name='"&CheckStr(Request.Form("gb_memName"))&"'"
		gb_AllreadyMem.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		IF gb_AllreadyMem.EOF AND gb_AllreadyMem.BOF Then
			gb_AllreadyMemErr=0
		ElseIF gb_AllreadyMem("mem_Password")=MD5(CheckStr(Request.Form("gb_MemPassword"))) Then
			Response.Cookies(CookieName)("memName")=gb_AllreadyMem("mem_Name")
			Response.Cookies(CookieName)("memPassword")=gb_AllreadyMem("mem_Password")
			Response.Cookies(CookieName)("memStatus")=gb_AllreadyMem("mem_Status")
			memName=gb_AllreadyMem("mem_Name")
			gb_AllreadyMem("mem_LastIP")=Guest_IP
			gb_AllreadyMem.Update
			gb_AllreadyMemErr=2
		Else
			gb_AllreadyMemErr=1
		End IF
		gb_AllreadyMem.Close
		Set gb_AllreadyMem=Nothing
		IF CheckStr(Request.Form("message"))=Empty OR CheckStr(Request.Form("gb_memName"))=Empty Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">请将必须信息填写完整，点击返回上一页</a>"
		ElseIF Len(CheckStr(Request.Form("gb_memName")))>24 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">用户名长度超过24个字符，12个汉字，点击返回上一页</a>"
		ElseIF Strurls(Request.Form("message"),"[url")>MaxUrl or Strurls(Request.Form("message"),"http://")>MaxHttp then 
                        msg_Title="错误信息"
                        msg_Content="<a href='javascript:history.go(-1);'>对不起，您输入的内容有非法链接，请返回重新输入</a>" 
		ElseIF IsValidUserName(CheckStr(Request.Form("gb_memName")))=False Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">用户名中含有非法字符，点击返回上一页</a>"
		ElseIF memName=Empty AND gb_AllreadyMemErr=1 Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">对不起，你所使用的用户名已经注册，点击返回上一页</a>"
		Else
			Dim gb_Content,gb_Title,gb_memName,gb_IsPublic
			gb_Content=CheckWordFilter(CheckStr(Request.Form("message")))
			gb_memName=CheckStr(Request.Form("gb_memName"))
			gb_IsPublic=Request.Form("gb_IsPublic")
			If Request.Form("gb_IsPublic")="1" Then
					gb_IsPublic=1
				Else
					gb_IsPublic=0
				End IF
			IF gb_IsPublic=Empty Then gb_IsPublic=0
			IF memName=Empty And gb_AllreadyMemErr<>2 Then
				Dim gb_SaveMem,gb_MemPassword
				gb_SaveMem=Request.Form("gb_SaveMem")
				gb_MemPassword=MD5(CheckStr(Request.Form("gb_MemPassword")))
				IF gb_SaveMem=1 Then
					Conn.ExeCute("INSERT INTO blog_Member(mem_Name,mem_Password,mem_LastIP) VALUES ('"&gb_memName&"','"&gb_memPassword&"','"&Guest_IP&"')")
					Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
					SQLQueryNums=SQLQueryNums+2
					Response.Cookies(CookieName)("memName")=gb_memName
					Response.Cookies(CookieName)("memPassword")=gb_memPassword
					Response.Cookies(CookieName)("memStatus")="Member"
				End IF
				Conn.ExeCute("INSERT INTO blog_Guestbook(gb_Content,gb_Author,gb_IsPublic,gb_PostIP) VALUES ('"&gb_Content&"','"&gb_Memname&"',"&gb_IsPublic&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			Else
				Conn.ExeCute("INSERT INTO blog_Guestbook(gb_Content,gb_Author,gb_IsPublic,gb_PostIP) VALUES ('"&gb_Content&"','"&memName&"',"&gb_IsPublic&",'"&Guest_IP&"')")
				SQLQueryNums=SQLQueryNums+1
			End IF
			Conn.ExeCute("UPDATE blog_Member SET mem_PostGBNums=mem_PostGBNums+1 WHERE mem_Name='"&gb_memName&"'")
			Conn.ExeCute("UPDATE blog_Info SET blog_GuestbookNums=blog_GuestbookNums+1")
			SQLQueryNums=SQLQueryNums+2
			Response.Cookies(CookieName)("memLastpost")=Now()
			msg_Title="发表成功"
			msg_Content="<a href='guestbook.asp'>留言发表成功，点击返回，或者3秒后自动返回</a><meta http-equiv='refresh' content='3;url=guestbook.asp'>"
		End If
	End If
	Response.Write("<br><br><center><div class=""msg_head"">"&msg_Title&"</div><div class=""msg_content"">"&msg_Content&"</div></center><br><br>")
ElseIf Request.QueryString("action")="delegb" Then
	IF IsInteger(Request.QueryString("gbID"))=False Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">参数出现错误，点击返回上一页</a>"
	Else
		IF Not (memStatus="SupAdmin" OR memStatus="Admin") Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">你没有权限删除评论，点击返回上一页</a>"
		Else
			Dim dele_GB
			Set dele_GB=Conn.ExeCute("SELECT gb_ID,gb_Author FROM blog_Guestbook WHERE gb_ID="&CheckStr(Request.QueryString("gbID")))
			SQLQueryNums=SQLQueryNums+1
			IF dele_GB.EOF AND dele_GB.BOF Then
				msg_Title="出现错误"
				msg_Content="<a href=""javascript:history.go(-1);"">没有找到指定留言，点击返回上一页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Info SET blog_GuestbookNums=blog_GuestbookNums-1")
				Conn.ExeCute("UPDATE blog_Member SET mem_PostGBNums=mem_PostGBNums-1 WHERE mem_Name='"&CheckStr(dele_GB("gb_Author"))&"'")
				Conn.Execute("DELETE  FROM blog_Guestbook WHERE gb_ID="&CheckStr(Request.QueryString("gbID")))
				SQLQueryNums=SQLQueryNums+3
				msg_Title="删除成功"
				msg_Content="<a href='guestbook.asp'>留言删除成功，点击返回</a>"
			End IF
			Set dele_GB=Nothing
		End If
	End IF
	Response.Write("<br><br><center><div class=""msg_head"">"&msg_Title&"</div><div class=""msg_content"">"&msg_Content&"</div></center><br><br>")
ElseIf Request.QueryString("action")="replygb" Then
	IF IsInteger(Request.QueryString("gbID"))=False Then
		msg_Title="出现错误"
		msg_Content="<a href=""javascript:history.go(-1);"">参数出现错误，点击返回上一页</a>"
	Else
		IF Not (memStatus="SupAdmin" OR memStatus="Admin") Then
			msg_Title="出现错误"
			msg_Content="<a href=""javascript:history.go(-1);"">你没有权限删除评论，点击返回上一页</a>"
		Else
			If CheckStr(Request.Form("message"))<>Empty Then
				Conn.ExeCute("UPDATE blog_Guestbook SET gb_Reply='"&CheckStr(Request.Form("message"))&"',gb_ReplyAuthor='"&memName&"',gb_ReplyTime='"&Now()&"' WHERE gb_ID="&Request.QueryString("gbID")&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="回复成功"
				msg_Content="<a href=""guestbook.asp"">点击返回留言本</a>"
			Else
			Dim reply_GB
			Set reply_GB=Conn.Execute("SELECT * FROM blog_Guestbook WHERE gb_ID="&Request.QueryString("gbID")&"")
			If reply_GB.Eof And reply_GB.Bof Then
				msg_Title="出现错误"
				msg_Content="<a href=""javascript:history.go(-1);"">你所评论的留言不存在，点击返回上一页</a>"
			Else
				msg_Title=reply_GB("gb_Author")&" 于 "&DateToStr(reply_GB("gb_PostTime"),"Y-m-d H:I A")&" 留言："
				msg_Content=Ubbcode(HTMLEncode(reply_GB("gb_Content")),0,0,0,1,0)
			%>
<script language="JavaScript" src="include/ubbcode.js"></script><table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td colspan="2" bgcolor="#EFEFEF"><b>回复留言</b></td>
        </tr><form name="input" method="post" action="guestbook.asp?action=replygb&gbID=<%=reply_GB("gb_ID")%>">
        <tr>
          <td valign="top" bgcolor="#FFFFFF" width="138"><b>回复内容：</b><div class="siderbar_head"><b>表&nbsp;&nbsp;情</b></div><div class="siderbar_main"><%Dim log_SmiliesListNums2,log_SmiliesListNumI2
			log_SmiliesListNums2=Ubound(Arr_Smilies,2)
			TempVar=""
			For log_SmiliesListNumI2=0 To log_SmiliesListNums2
				Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI2)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI2)&""" />")
				TempVar=" "
			Next
			%></div></td>
          <td bgcolor="#FFFFFF" nowrap><textarea name="message" cols="62" rows="8" wrap="VIRTUAL" id="Message" onSelect="javascript: storeCaret(this);" onClick="javascript: storeCaret(this);" onKeyUp="javascript: storeCaret(this);" onKeyDown="javascript: ctlent();"><%=EditDeHTML(reply_GB("gb_Reply"))%></textarea></td>
        </tr>
        <tr align="center">
          <td colspan="2" nowrap bgcolor="#FFFFFF"><input type="submit" name="replysubmit" value=" 回复留言 [可按 Ctrl+Enter 发布] " onClick="this.disabled=true;document.input.submit();" />
&nbsp; <input name="Reset" type="reset" id="Reset" value=" 重置回复 " /></td>
          </tr></form>
      </table>
<%End If
Set reply_GB=Nothing
End If
End If
End If
Response.Write("<br><br><center><div class=""msg_head"">"&msg_Title&"</div><div class=""msg_content"">"&msg_Content&"</div></center><br><br>")
Else
	Dim CurPage
	If CheckStr(Request.QueryString("Page"))<>Empty Then
		Curpage=CheckStr(Request.QueryString("Page"))
		If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
	Else
		Curpage=1
	End If

	Dim RS_GB
	Set RS_GB=Server.CreateObject("Adodb.Recordset")
	SQL="SELECT * FROM blog_Guestbook "&SQLFiltrate&" ORDER BY GB_ID DESC"
	RS_GB.Open SQL,CONN,1,1
	SQLQueryNums=SQLQueryNums+1
	If RS_GB.EOF AND RS_GB.BOF Then 
		Response.Write("<div class=""message"">暂时没有留言</div>")
	Else
		Dim GB_Nums,gbview_ID,gbview_Author,gbview_Reply
		RS_GB.PageSize=blogPerPage
		RS_GB.AbsolutePage=CurPage
		GB_Nums=RS_GB.RecordCount
		Dim MultiPages,PageCount
		MultiPages="<span class=""smalltxt"">"&MultiPage(GB_Nums,blogPerPage,CurPage,Url_Add)&"</span>"
		Response.Write(MultiPages)
		Response.Write("<table width=""100%"" height=""6"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
		Do Until RS_GB.EOF OR PageCount=blogPerPage
			gbview_ID=RS_GB("gb_ID")
			gbview_Author=RS_GB("gb_Author")
			gbview_Reply=RS_GB("gb_ReplyAuthor")
			Response.Write("<div class=""content_head"" style=""font-size:11px;""><b>")
			If gbview_Reply<>Empty Then
				Response.Write("<img src=""images/icon_ctb.gif"" border=""0"" align=""absmiddle""/>")
			ElseIf RS_GB("gb_IsPublic")=0 Then
				Response.Write("<img src=""images/icon_cttb.gif"" border=""0"" align=""absmiddle""/>")
			Else
				Response.Write("<img src=""images/icon_quote.gif"" border=""0"" align=""absmiddle""/>")
			End If
			Response.Write("&nbsp;<a href=""member.asp?action=view&memName="&Server.URLEncode(gbview_Author)&""" target=""_blank"">"&gbview_Author&"</a> 于 "&DateToStr(RS_GB("gb_PostTime"),"Y-m-d H:I A")&" 留言：</b>")
			If memStatus="SupAdmin" OR memStatus="Admin" Then
				Response.Write("&nbsp;&nbsp;<a href=""guestbook.asp?action=delegb&gbID="&gbview_ID&""" title=""删除留言"" onClick=""winconfirm('你真的要删除这个留言吗？','guestbook.asp?action=delegb&gbID="&gbview_ID&"'); return false""><b><font color=""#FF0000"">×</font></b></a>&nbsp;|&nbsp;<a href=""ipview.asp?ipdata="&RS_GB("gb_PostIP")&""" title=""点击查看IP地址："&RS_GB("gb_PostIP")&" 的来源"" target=""_blank"">IP</a>&nbsp;|&nbsp;<a href=""guestbook.asp?action=replygb&gbID="&gbview_ID&""" title=""点击回复"">回复</a>")
			End If
			Response.Write("</div><div class=""content_main"" style=""font-size:11px;"">")
			If RS_GB("gb_IsPublic")=1 OR (memStatus="SupAdmin" OR memStatus="Admin" OR memName=RS_GB("gb_Author")) Then
				Response.Write(Ubbcode(HTMLEncode(RS_GB("gb_Content")),0,0,0,1,0))
				If gbview_Reply<>Empty Then
					Response.Write("<div class=""content_head"" style=""margin-left:12px;margin-right:12px;margin-top:5px;""><span class=""smalltxt""><b><img src=""images/icon_ctb.gif"" border=""0"" align=""absmiddle""/> <a href=""member.asp?action=view&memName="&Server.URLEncode(RS_GB("gb_ReplyAuthor"))&""" target=""_blank"">"&RS_GB("gb_ReplyAuthor")&"</a> 于 "&DateToStr(RS_GB("gb_ReplyTime"),"Y-m-d H:I A")&" 回复：</b><div style=""margin-left:6px;margin-top:6px;"">"&Ubbcode(HTMLEncode(RS_GB("gb_Reply")),0,0,0,1,0)&"</div></span></div>")
				End If
			Else
				Response.Write("此留言仅管理员以及发表人可以浏览")
			End If
			Response.Write("</div>")
			RS_GB.MoveNext
			PageCount=PageCount+1
			Response.Write("<table width=""100%"" height=""6"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")
		Loop
	End If
	RS_GB.Close
	Set RS_GB=Nothing
	Response.Write(MultiPages)
	Response.Write("<table width=""100%"" height=""6"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td></td></tr></table>")%>
<script language="JavaScript" src="include/ubbcode.js"></script><table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
        <tr>
          <td colspan="2" bgcolor="#EFEFEF"><b>发表留言</b><b>发表留言 - 不要忘了输入验证码哦！</b></td>
        </tr><form name="input" method="post" action="guestbook.asp?action=postgb">
        <tr>
          <td width="82" align="right" bgcolor="#FFFFFF" nowrap><b>作者：</b></td>
          <td width="100%" bgcolor="#FFFFFF">用户名：
            <%IF memName<>Empty Then
			Response.Write("<input name=""gb_memName"" type=""text"" id=""gb_memName"" value="""&memName&""" size=""10"" readonly />")
			Else
			Response.Write("<input name=""gb_memName"" type=""text"" id=""gb_memName"" size=""8"" />&nbsp;密码： <input name=""gb_memPassword"" type=""password"" id=""gb_memPassword"" size=""8"" />&nbsp;<input name=""gb_SaveMem"" type=""checkbox"" id=""gb_SaveMem"" value=""1"" /> 同时注册？")
			End IF%>&nbsp;验证：<input name="validatecode" type="text" id="validatecode" size="3" />&nbsp;<img src="include/validatecode.asp" align="absmiddle" border="0" /></td>
        </tr>
        <tr>
          <td align="right" valign="top" bgcolor="#FFFFFF"><b>内容：
          </b><div style="padding-left:5px;" align="left" width="100%">
                <br><input name="gb_IsPublic" type="checkbox" id="gb_IsPublic" value="1" checked/>公开留言</div></td>
          <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td valign="top" nowrap><textarea name="message" cols="85" rows="8" wrap="VIRTUAL" id="Message" onSelect="javascript: storeCaret(this);" onClick="javascript: storeCaret(this);" onKeyUp="javascript: storeCaret(this);" onKeyDown="javascript: ctlent();"></textarea></td>
                <td width="122" align="left" valign="top"><div class="siderbar_head"><b>表&nbsp;&nbsp;情</b></div><div class="siderbar_main"><%Dim log_SmiliesListNums,log_SmiliesListNumI
			log_SmiliesListNums=Ubound(Arr_Smilies,2)
			TempVar=""
			For log_SmiliesListNumI=0 To log_SmiliesListNums
				Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" />")
				TempVar=" "
			Next
			%></div></td>
              </tr>
            </table></td>
        </tr>
        <tr align="center">
          <td colspan="2" nowrap bgcolor="#FFFFFF"><input type="submit" name="replysubmit" value=" 发表留言 [可按 Ctrl+Enter 发布] " onClick="this.disabled=true;document.input.submit();" />
&nbsp; <input name="Reset" type="reset" id="Reset" value=" 重置留言 " /></td>
          </tr></form>
      </table><%End If%>
</td>
 </tr>
</table><!--#include file="footer.asp" -->