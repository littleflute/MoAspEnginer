<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<%Dim ForumID,ThreadID,Url_Add,forum_Title,found_Error
	ForumID=Trim(Request.QueryString("forumID"))
	ThreadID=Trim(Request.QueryString("threadID"))
	found_Error=False
	Url_Add="?"
	If IsInteger(ForumID)=True Then
		Url_Add=Url_Add&"forumID="&ForumID&"&"
	End If
	If IsInteger(ThreadID)=True Then
		Url_Add=Url_Add&"threadID="&ThreadID&"&"
	End If%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call ForumList(ForumID)
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call forumSearch
	%><br>
	</td>
	<td width="100%" valign="top" bgcolor="#FFFFFF"><%
	If IsInteger(ForumID)=False OR IsInteger(ThreadID)=False Then
		Response.Write("<div class=""message""><a href=""javascript:history.go(-1)"">对不起，参数错误，点击返回重新操作</a></div>")
	Else
		Dim CurPage
	    Call GetCurPage
		Dim ThreadView
		Set ThreadView=Server.CreateObject("Adodb.Recordset")
		SQL="SELECT M.mem_Intro,T.thread_Title,T.thread_Icon,T.thread_PostTime,T.thread_IsTop,T.thread_IsDigest,T.thread_IsClose,T.thread_MagicFace,P.post_ForumID,P.post_ThreadID,P.post_Content,P.post_IsClosed,P.post_DisSM,P.post_DisIMG,P.post_DisUBB,P.post_AutoURL,P.post_AutoKEY,P.post_ID,P.post_IsTop,P.post_Author,P.post_PostTime,P.post_PostIP,P.post_Modify FROM blog_Threads AS T,blog_Posts AS P,blog_Member AS M WHERE T.thread_ID="&ThreadID&" AND P.post_ThreadID="&ThreadID&" AND M.mem_Name=P.post_Author ORDER BY P.post_IsTop,P.post_ID ASC"
		ThreadView.Open SQL,CONN,1,1
		SQLQueryNums=SQLQueryNums+1
		If ThreadView.EOF And ThreadView.BOF Then
			Response.Write("<div class=""message""><a href=""javascript:history.go(-1)"">对不起，找不到相关主题，点击返回重新操作</a></div>")
		Else
			Dim Post_Num,post_forumID,post_ID,post_Author,post_Floors,post_threadID,thread_IsTop,thread_IsDigest,thread_IsClose,thread_MagicFace,thread_MagicFaceStr,post_IsTop
			post_forumID=ThreadView("post_ForumID")
			post_threadID=ThreadView("post_ThreadID")
			thread_IsTop=ThreadView("thread_IsTop")
			thread_IsDigest=ThreadView("thread_IsDigest")
			thread_IsClose=ThreadView("thread_IsClose")
			thread_MagicFace=ThreadView("thread_MagicFace")
			Conn.ExeCute("UPDATE blog_Threads SET thread_ViewNums=thread_ViewNums+1 WHERE thread_ID="&ThreadID&"")
			SQLQueryNums=SQLQueryNums+1
			forum_Title=ForumTitle(post_forumID)&" - "&HTMLEncode(ThreadView("thread_Title"))
			Response.Write("<div class=""content_head""><a href=""forumview.asp"">Acblog BBS</a>"&forum_Title&"</div><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table>")
			ThreadView.PageSize=postPerPage
			ThreadView.AbsolutePage=CurPage
			Post_Num=ThreadView.RecordCount
			Dim MultiPages,PageCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(Post_Num,postPerPage,CurPage,Url_Add)&"</span>"
			Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"&MultiPages&"</td><td align=""right""><a href=""forumpost.asp?action=thread&forumID="&ThreadView("post_ForumID")&"""><img src=""images/newthread.gif"" border=""0"" align=""absmiddle""></a><a href=""forumpost.asp?action=reply&forumID="&ThreadView("post_ForumID")&"&threadID="&ThreadView("post_ThreadID")&"""><img src=""images/newreply.gif"" border=""0"" align=""absmiddle""></a></td></tr></table><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table><table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" class=""wordbreak"">")
			Do Until ThreadView.EOF Or PageCount=postPerPage
				post_Floors=(CurPage-1)*postPerPage+PageCount
				If post_Floors=0 Then
					post_Floors="<b>楼&nbsp;&nbsp;主</b>"
				Else
					post_Floors="第 "&post_Floors+1&" 楼"
				End If
				post_Author=ThreadView("post_Author")
				post_ID=ThreadView("post_ID")
				post_IsTop=ThreadView("post_IsTop")
				Response.Write("<tbody><tr><td colspan=""2""><table class=""content_head"" cellpadding=""4"" cellspacing=""0""><tr class=""smalltxt""><td width=""100%""><b><a name=""post_"&post_ID&"""></a><a href=""forumpost.asp?action=reply&forumID="&post_forumID&"&threadID="&post_threadID&"&quoteID="&post_ID&"""><image src=""images/icon_quote.gif"" border=""0"" align=""absmiddle"" alt=""引用回复""></a>&nbsp;<a href=""member.asp?action=view&memName="&Server.URLEncode(post_Author)&""">"&post_Author&"</a> 于 "&DateToStr(ThreadView("post_PostTime"),"Y-m-d H:I A")&"</b> 发表：")
				If memStatus="SupAdmin" Then
					If post_IsTop=False Then Response.Write("&nbsp;<a href=""forummisc.asp?action=delepost&forumID="&post_forumID&"&threadID="&post_threadID&"&postID="&post_ID&""" title=""删除帖子"" onClick=""winconfirm('你真的要删除这个帖子吗？','forummisc.asp?action=delepost&forumID="&post_forumID&"&threadID="&post_threadID&"&postID="&post_ID&"'); return false""><b><font color=""#FF0000"">×</font></b></a>&nbsp;|")
					Response.Write("&nbsp;<a href="""&IPViewURL&ThreadView("post_PostIP")&""" title=""点击查看IP地址："&ThreadView("post_PostIP")&" 的来源"" target=""_blank"">IP</a>&nbsp;|&nbsp;<a href=""forummisc.asp?action=closepost&forumID="&post_forumID&"&threadID="&post_threadID&"&postID="&post_ID&""" title=""屏蔽帖子"" onClick=""winconfirm('你真的要屏蔽这个帖子吗？','forummisc.asp?action=closepost&forumID="&post_forumID&"&threadID="&post_threadID&"&postID="&post_ID&"'); return false""><img src=""images/icon_postclose.gif"" align=""absmiddle"" border=""0""></a>")
				End If
				If post_Author=memName And (memStatus="SupAdmin" Or DateDiff("d",ThreadView("post_PostTime"),Now())<3) Then
					If post_IsTop=True Then
						Response.Write("&nbsp;|&nbsp;<a href=""forumedit.asp?action=thread&forumID="&post_forumID&"&threadID="&post_threadID&"""><img src=""images/icon_edit.gIf"" border=""0"" align=""absmiddle"" alt=""编辑主题""></a>")
					Else
						Response.Write("&nbsp;|&nbsp;<a href=""forumedit.asp?action=reply&forumID="&post_forumID&"&threadID="&post_threadID&"&postID="&post_ID&"""><img src=""images/icon_edit.gIf"" border=""0"" align=""absmiddle"" alt=""编辑回复""></a>")
					End If
				End If
				If post_IsTop=True And thread_MagicFace<>Empty Then 
					thread_MagicFaceStr="<div id=""MagicFace"" style=""position:absolute;z-index:99;visibility:hidden;""></div><script type=""text/javascript"" src=""include/magicface.js""></script><script language=""JavaScript"">ShowMagicFace('"&thread_MagicFace&"');</script><div class=""magicface""><img src=""magicface/images/"&thread_MagicFace&".gif"" align=""absmiddle""title=""点击观看魔法表情"" style=""cursor:hand;"" onClick=""ShowMagicFace('"&thread_MagicFace&"');""><br>魔法表情</div>"
				Else
					thread_MagicFaceStr=""
				End If
				Response.Write("</td><td nowrap>"&post_Floors&"</td></tr></table><div class=""content_main"">")
				If ThreadView("post_IsClosed")=False Then 
					Response.Write(thread_MagicFaceStr&"<div style=""clear:left;"">"&Ubbcode(HTMLEncode(ThreadView("post_Content")),ThreadView("post_DisSM"),ThreadView("post_DisUBB"),ThreadView("post_DisIMG"),ThreadView("post_AutoURL"),ThreadView("post_AutoKEY"))&"</div>")
					If  ThreadView("post_Modify")<>Empty Then
							Response.Write("<div align=""right"" class=""smalltxt"" height=""35px"">"&ThreadView("post_Modify")&"</div>")
			End IF
				Else
					Response.Write("<br><center><b>此帖已被管理员屏蔽</b></center><br>")
					If memStatus="SupAdmin" Or memStatus="Admin" Then
						Response.Write("<div class=""smalltxt""><span class=""siderbar_main"" style=""cursor:hand;"" onclick=""showIntro('adminshowpost_"&post_ID&"');""><img src=""images/icon_postclose.gif"" align=""absmiddle"" border=""0"">&nbsp;<b>管理员查看帖子内容：</b></span><br><div class=""content_main"" style=""display:none;"" id=""adminshowpost_"&post_ID&""">"&Ubbcode(HTMLEncode(ThreadView("post_Content")),ThreadView("post_DisSM"),ThreadView("post_DisUBB"),ThreadView("post_DisIMG"),ThreadView("post_AutoURL"),ThreadView("post_AutoKEY"))&"")
						If ThreadView("post_Modify")<>Empty Then
							Response.Write("<div align=""right"" class=""smalltxt"" height=""35px"">"&ThreadView("post_Modify")&"</div>")
						End If
						Response.Write("</div></div>")
					End If
				End If
				Response.Write("</div>")
				If ThreadView("mem_Intro")<>Empty Then Response.Write("<div class=""content_main""><img src=""images/sigline.gif""><br />"&Realremark(UbbCode(HTMLEncode(ThreadView("mem_Intro")),1,0,1,0,0))&"</div>")
				Response.Write("</td></tr></tbody>")
				ThreadView.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("</table><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table>")
			Response.Write(MultiPages)
			If (memName<>Empty And thread_IsClose=False) Or (memStatus="SupAdmin" Or memStatus="Admin") Then
				Response.Write("<div class=""content_main""><script language=""JavaScript"" src=""include/ubbcode.js""></script><table width=""100%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr><td colspan=""2"" bgcolor=""#EFEFEF""><b>快速回复&nbsp;-&nbsp;[本站已启用NoFollow标签]</b></td></tr><form name=""input"" method=""post"" action=""forumpost.asp?action=reply&forumID="&post_forumID&"&threadID="&post_threadID&"""><tr><td align=""right"" valign=""top"" bgcolor=""#FFFFFF""><b>内容：</b><div style=""padding-left:5px;"" align=""left"" width=""100%""><br><input name=""post_DisSM"" type=""checkbox"" id=""post_DisSM"" value=""1"" /> 禁止表情<br><input name=""post_DisUBB"" type=""checkbox"" id=""post_DisUBB"" value=""1"" /> 禁止UBB<br><input name=""post_DisIMG"" type=""checkbox"" id=""post_DisIMG"" value=""1"" /> 禁止图片<br><input name=""post_AutoURL"" type=""checkbox"" id=""post_AutoURL"" value=""1"" checked /> 识别链接<br><input name=""post_AutoKEY"" type=""checkbox"" id=""post_AutoKEY"" value=""1"" /> 识别关键字</div></td><td bgcolor=""#FFFFFF""><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td valign=""top"" nowrap><textarea name=""message"" cols=""50"" rows=""8"" wrap=""VIRTUAL"" id=""Message"" onSelect=""javascript: storeCaret(this);"" onClick=""javascript: storeCaret(this);"" onKeyUp=""javascript: storeCaret(this);"" onKeyDown=""javascript: ctlent();""></textarea></td><td width=""122"" align=""left"" valign=""top""><div class=""siderbar_head""><b>表&nbsp;&nbsp;情</b></div><div class=""siderbar_main"">")
				Dim log_SmiliesListNums,log_SmiliesListNumI
				log_SmiliesListNums=Ubound(Arr_Smilies,2)
				TempVar=""
				For log_SmiliesListNumI=0 To log_SmiliesListNums
					Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&""" />")
					TempVar=" "
				Next
				Response.Write("</div></td></tr></table><iframe border=""0"" frameBorder=""0"" frameSpacing=""0"" height=""28"" marginHeight=""0"" marginWidth=""0"" noResize scrolling=""no"" width=""100%"" vspale=""0"" src=""attachment.asp""></iframe></td></tr><tr align=""center""><td colspan=""2"" nowrap bgcolor=""#FFFFFF""><input name=""IsPostOK"" type=""hidden"" value=""PostIT""><input type=""submit"" name=""topicsubmit"" value="" 发表回复 [可按 Ctrl+Enter 发布] "" onClick=""this.disabled=true;document.input.submit();"" />&nbsp; <input name=""Reset"" type=""reset"" id=""Reset"" value="" 重置回复 "" /></td></tr></form></table></div>")
			End If
			If memStatus="SupAdmin" Then
				Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table><div class=""content_head""><span class=""smalltxt""><b>管理选项：</b>&nbsp;<a href=""forummisc.asp?action=delethread&forumID="&post_forumID&"&threadID="&post_threadID&""" onClick=""winconfirm('你真的要删除这个主题吗？','forummisc.asp?action=delethread&forumID="&post_forumID&"&threadID="&post_threadID&"'); return false"">删除主题</a>&nbsp;&nbsp;|&nbsp;&nbsp;")
				If thread_IsTop=0 Then
					Response.Write("<a href=""forummisc.asp?action=top&threadID="&post_threadID&""">置顶主题</a>")
				Else
					Response.Write("<a href=""forummisc.asp?action=untop&threadID="&post_threadID&""">取消置顶</a>")
				End If
				Response.Write("&nbsp;&nbsp;|&nbsp;&nbsp;")
				If thread_IsDigest=False Then
					Response.Write("<a href=""forummisc.asp?action=digest&threadID="&post_threadID&""">加入精华</a>")
				Else
					Response.Write("<a href=""forummisc.asp?action=undigest&threadID="&post_threadID&""">取消精华</a>")
				End If
				Response.Write("&nbsp;&nbsp;|&nbsp;&nbsp;")
				If thread_IsClose=False Then
					Response.Write("<a href=""forummisc.asp?action=close&threadID="&post_threadID&""">锁定主题</a>")
				Else
					Response.Write("<a href=""forummisc.asp?action=unclose&threadID="&post_threadID&""">解除锁定</a>")
				End If
				Response.Write("</span></div>")
			End If
		End If
		ThreadView.Close
		Set ThreadView=Nothing
	End If
	%></td>
  </tr>
</table>
<!--#include file="footer.asp" -->