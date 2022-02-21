<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
	  <tr><td align="center" valign="top">
<%
If memName<>Empty Then
	If (memStatus<>"SupAdmin" AND memStatus<>"Admin") And DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<15 Then
		Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href='javascript:history.go(-1);'>你发表的速度太快了，点击返回上一页</a><br><br><br></div>")
	Else
		Dim post_Action,forum_ID,thread_ID,quote_ID
		post_Action=Trim(Request.QueryString("action"))
		forum_ID=Trim(Request.QueryString("forumID"))
		If forum_ID=Empty Then forum_ID=Trim(Request.Form("forumID"))
		thread_ID=Trim(Request.QueryString("threadID"))
		quote_ID=Trim(Request.QueryString("quoteID"))
		If post_Action="thread" Then 
			If IsInteger(forum_ID)=False Then
				Call SelectForum
			Else
				Dim SelectedForumRS
				Set SelectedForumRS=Conn.ExeCute("SELECT forum_Name FROM blog_Forums WHERE forum_ID="&forum_ID&"")
				SQLQueryNums=SQLQueryNums+1
				If SelectedForumRS.EOF And SelectedForumRS.BOF Then
					Call SelectForum
				Else
					If Request.Form("IsPostOK")<>Empty Then
						Call ForumPostSave(post_Action,forum_ID,Empty)
					Else
						Call ForumPost(post_Action,forum_ID,SelectedForumRS("forum_Name"),Empty,Empty,Empty)
					End If
				End If
				Set SelectedForumRS=Nothing
			End If
		ElseIf post_Action="reply" Then
			If IsInteger(forum_ID)=False OR IsInteger(thread_ID)=False Then
				Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">参数无效，请点击返回上一页！</a><br><br><br></div>")
			Else
				Dim ReplyThreadRS
				Set ReplyThreadRS=Conn.ExeCute("SELECT thread_Title FROM blog_Threads WHERE thread_ID="&thread_ID&" AND thread_ForumID="&forum_ID&"")
				SQLQueryNums=SQLQueryNums+1
				If ReplyThreadRS.EOF And ReplyThreadRS.BOF Then
					Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">你所回复的主题不存在，请点击返回上一页！</a><br><br><br></div>")
				Else
					If Request.Form("IsPostOK")<>Empty Then
						Call ForumPostSave(post_Action,forum_ID,thread_ID)
					Else
						Call ForumPost(post_Action,forum_ID,Empty,thread_ID,ReplyThreadRS("thread_Title"),quote_ID)
					End If
				End If
				Set ReplyThreadRS=Nothing
			End If
		Else
			Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">参数无效，请点击返回上一页！</a><br><br><br></div>")
		End If
	End If
Else
	Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""logging.asp"">游客不能发表帖子，请先登陆</a><br><br><br></div>")
End If

Sub ForumPost(Action,ForumID,ForumName,ThreadID,ThreadTitle,QuoteID)
	Dim TitleContent
	Response.Write("<table width=""97%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr align=""center""><td colspan=""3"" class=""msg_head""><script language=""JavaScript"" src=""include/ubbcode.js""></script><script language=""JavaScript"" src=""include/ubbhelp.js""></script>")
	If Action="thread" Then
		Response.Write("在论坛&nbsp;<font color=""#FFFA20"">"&ForumName&"</font>&nbsp;中发表新主题")
		TitleContent="<tr bgcolor=""#FFFFFF""><td width=""130"" align=""right"" nowrap><b>标题：</b></td><td width=""100%""><input name=""post_Title"" type=""text"" id=""post_Title"" size=""50"">   <select name=""TitleFontColor"" id=""TitleFontColor""><option value="""" selected>颜色</option><OPTION value="""">默认</OPTION><OPTION value=""#000000"" style=""background-color:#000000""></OPTION><OPTION value=""#FFFF00"" style=""background-color:#FFFF00""></OPTION><OPTION value=""#00FF00"" style=""background-color:#00FF00""></OPTION><OPTION value=""#FF00FF"" style=""background-color:#FF00FF""></OPTION><OPTION value=""#FF0000"" style=""background-color:#FF0000""></OPTION><OPTION value=""#0000FF"" style=""background-color:#0000FF""></OPTION></select></td></tr><tr bgcolor=""#FFFFFF""><td align=""right"" nowrap><b>图标：</b></td><td width=""100%""><input type=""radio"" name=""post_Icon"" value="""" checked> 无  <input type=""radio"" name=""post_Icon"" value=""icon1.gif""><img src=""images/threadicon/icon1.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon2.gif""><img src=""images/threadicon/icon2.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon3.gif""><img src=""images/threadicon/icon3.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon4.gif""><img src=""images/threadicon/icon4.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon5.gif""><img src=""images/threadicon/icon5.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon6.gif""><img src=""images/threadicon/icon6.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon7.gif""><img src=""images/threadicon/icon7.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon8.gif""><img src=""images/threadicon/icon8.gif"" align=""absmiddle""> <input type=""radio"" name=""post_Icon"" value=""icon9.gif""><img src=""images/threadicon/icon9.gif"" align=""absmiddle""></td></tr>"
		If memStatus="SupAdmin" Or memStatus="Admin" Then TitleContent=TitleContent&"<tr bgcolor=""#FFFFFF""><td align=""right""><strong>属性：</strong></td><td colspan=""2""><input name=""Post_IsTop"" type=""checkbox"" id=""post_IsTop"" value=""1""> 置顶主题&nbsp;&nbsp;<input name=""post_IsDigest"" type=""checkbox"" id=""post_IsDigest"" value=""1""> 精华主题&nbsp;&nbsp;<input name=""post_IsClose"" type=""checkbox"" id=""post_IsClose"" value=""1""> 锁定主题&nbsp;&nbsp;<input name=""post_IsPoll"" type=""checkbox"" id=""post_IsPoll"" value=""1""> 投票主题</td></tr>"
	ElseIf Action="reply" Then
		Response.Write("回复主题")
		TitleContent="<tr bgcolor=""#FFFFFF""><td width=""130"" align=""right"" nowrap><b>主题：</b></td><td width=""100%"">"&EditDeHTML(ThreadTitle)&"</td></tr>"
	End If
	Response.Write("</td></tr>")
	If Action="thread" Then
		Response.Write("<form name=""input"" method=""post"" action=""?action="&Action&"&forumID="&ForumID&""">")
	ElseIf Action="reply" Then
		Response.Write("<form name=""input"" method=""post"" action=""?action="&Action&"&forumID="&ForumID&"&threadID="&ThreadID&""">")
	End If
	Response.Write(TitleContent&"<tr bgcolor=""#FFFFFF""><td align=""right"" valign=""top"" class=""wordbreak""><b>内容：<br><br></b><div class=""siderbar_head"">表&nbsp;&nbsp;情</div><div class=""siderbar_main"" align=""left"">")
	Dim log_SmiliesListNums,log_SmiliesListNumI
	log_SmiliesListNums=Ubound(Arr_Smilies,2)
	TempVar=""
	For log_SmiliesListNumI=0 To log_SmiliesListNums
		Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&"""/>")
		TempVar=" "
	Next
	Response.Write("</div><div style=""padding-left:15px;"" align=""left"" width=""100%""><br><br><input name=""post_DisSM"" type=""checkbox"" id=""post_DisSM"" value=""1""> 禁止表情<br><input name=""post_DisUBB"" type=""checkbox"" id=""post_DisUBB"" value=""1""> 禁止UBB<br><input name=""post_DisIMG"" type=""checkbox"" id=""post_DisIMG"" value=""1""> 禁止图片<br><input name=""post_AutoURL"" type=""checkbox"" id=""post_AutoURL"" value=""1"" checked> 识别链接<br><input name=""post_AutoKEY"" type=""checkbox"" id=""post_AutoKEY"" value=""1""> 识别关键字</div></td><td colspan=""2""><table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td width=""100%""><select name=""font"" onFocus=""this.selectedIndex=0"" onChange=""chfont(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>选择字体</option><option value=""宋体"">宋体</option><option value=""黑体"">黑体</option><option value=""Arial"">Arial</option><option value=""Book Antiqua"">Book Antiqua</option><option value=""Century Gothic"">Century Gothic</option><option value=""Courier New"">Courier New</option><option value=""Georgia"">Georgia</option><option value=""Impact"">Impact</option><option value=""Tahoma"">Tahoma</option><option value=""Times New Roman"">Times New Roman</option><option value=""Verdana"">Verdana</option></select><select name=""size"" onFocus=""this.selectedIndex=0"" onChange=""chsize(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>字体大小</option><option value=""-2"">-2</option><option value=""-1"">-1</option><option value=""1"">1</option><option value=""2"">2</option><option value=""3"">3</option><option value=""4"">4</option><option value=""5"">5</option><option value=""6"">6</option><option value=""7"">7</option></select><select name=""color""  onFocus=""this.selectedIndex=0"" onChange=""chcolor(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>字体颜色</option><option value=""White"" style=""background-color:white;color:white;"">White</option><option value=""Black"" style=""background-color:black;color:black;"">Black</option><option value=""Red"" style=""background-color:red;color:red;"">Red</option><option value=""Yellow"" style=""background-color:yellow;color:yellow;"">Yellow</option><option value=""Pink"" style=""background-color:pink;color:pink;"">Pink</option><option value=""Green"" style=""background-color:green;color:green;"">Green</option><option value=""Orange"" style=""background-color:orange;color:orange;"">Orange</option><option value=""Purple"" style=""background-color:purple;color:purple;"">Purple</option><option value=""Blue"" style=""background-color:blue;color:blue;"">Blue</option><option value=""Beige"" style=""background-color:beige;color:beige;"">Beige</option><option value=""Brown"" style=""background-color:brown;color:brown;"">Brown</option><option value=""Teal"" style=""background-color:teal;color:teal;"">Teal</option><option value=""Navy"" style=""background-color:navy;color:navy;"">Navy</option><option value=""Maroon"" style=""background-color:maroon;color:maroon;"">Maroon</option><option value=""LimeGreen"" style=""background-color:limegreen;color:limegreen;"">LimeGreen</option></select></td><td rowspan=""2"" nowrap><b>模式：</b>&nbsp;<input type=""radio"" name=""mode"" value=""2"" onclick=""chmode('2')"" checked> 基本&nbsp;<input type=""radio"" name=""mode"" value=""0"" onclick=""chmode('0')""> 高级&nbsp;<input type=""radio"" name=""mode"" value=""1"" onclick=""chmode('1')""> 帮助</td></tr><tr><td><a href=""javascript:bold()""><img src=""images/ubbcode/bb_bold.gif"" border=""0"" alt=""插入粗体文本""></a> <a href=""javascript:italicize()""><img src=""images/ubbcode/bb_italicize.gif"" border=""0"" alt=""插入斜体文本""></a> <a href=""javascript:underline()""><img src=""images/ubbcode/bb_underline.gif"" border=""0"" alt=""插入下划线""></a> <a href=""javascript:center()""><img src=""images/ubbcode/bb_center.gif"" border=""0"" alt=""居中对齐""></a> <a href=""javascript:hyperlink()""><img src=""images/ubbcode/bb_url.gif"" border=""0"" alt=""插入超级链接""></a> <a href=""javascript:email()""><img src=""images/ubbcode/bb_email.gif"" border=""0"" alt=""插入邮件地址""></a> <a href=""javascript:image()""><img src=""images/ubbcode/bb_image.gif"" border=""0"" alt=""插入图像""></a> <a href=""javascript:flash()""><img src=""images/ubbcode/bb_flash.gif"" border=""0"" alt=""插入 Flash""></a> <a href=""javascript:code()""><img src=""images/ubbcode/bb_code.gif"" border=""0"" alt=""插入代码""></a> <a href=""javascript:quote()""><img src=""images/ubbcode/bb_quote.gif"" border=""0"" alt=""插入引用""></a> <a href=""javascript:list()""><img src=""images/ubbcode/bb_list.gif"" border=""0"" alt=""插入列表""></a> <a href=""javascript:wma()""><img src=""images/ubbcode/bb_wma.gif"" alt=""插入音频文件"" width=""23"" height=""22"" border=""0""></a> <a href=""javascript:wmv()""><img src=""images/ubbcode/bb_wmv.gif"" alt=""插入视频文件"" width=""23"" height=""22"" border=""0""></a></td></tr></table><table width=""98%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr valign=""top""><td><textarea name=""message"" style=""width:100%"" rows=""18"" wrap=""VIRTUAL"" id=""Message"" onSelect=""javascript: storeCaret(this);"" onClick=""javascript: storeCaret(this);"" onKeyDown=""javascript: ctlent();"" onKeyUp=""javascript: storeCaret(this);"">")
	If Action="reply" And IsInteger(QuoteID)=True Then
		Dim ReplyQuoteRS
		Set ReplyQuoteRS=Conn.ExeCute("SELECT post_Content,post_Author,post_PostTime FROM blog_Posts WHERE post_ID="&QuoteID&"")
		SQLQueryNums=SQLQueryNums+1
		If Not(ReplyQuoteRS.Eof AND ReplyQuoteRS.Bof) Then
			Response.Write("[quote][b]最初由 [color=blue]"&ReplyQuoteRS("post_Author")&"[/color] 发表于 "&DateToStr(ReplyQuoteRS("post_PostTime"),"Y-m-d H:I A")&"：[/b]"&CHR(10)&Replace(cutStr(DelQuote(Replace(HTMLEncode(ReplyQuoteRS("post_Content")),"<br><br>","<br>")),80),"<br>",CHR(10))&"...[/quote]"&CHR(10))
		End If
	End If
	Response.Write("</textarea></td></tr></table></td></tr><tr bgcolor=""#FFFFFF""><td align=""right""><b>附件：</b></td><td colspan=""2""><iframe border=""0"" frameBorder=""0"" frameSpacing=""0"" height=""28"" marginHeight=""0"" marginWidth=""0"" noResize scrolling=""no"" width=""100%"" vspale=""0"" src=""attachment.asp""></iframe></td></tr><tr align=""center"" bgcolor=""#FFFFFF""><td colspan=""3""><input name=""IsPostOK"" type=""hidden"" value=""PostIT""><input name=""topicsubmit"" type=""submit"" value="" 提交帖子 [可按 Ctrl+Enter 发布] "" onClick=""this.disabled=true;document.input.submit();"">&nbsp;<input name=""L_Reset"" type=""reset"" id=""L_Reset"" value="" 重置帖子 ""></td></tr></form></table>")
	If Action="reply" Then
		Dim R_ViewThread,R_ViewThreadCount,R_ViewThreadContent
		Set R_ViewThread=Server.CreateObject("ADODB.Recordset")
		SQL="SELECT TOP "&postPerPage+1&" post_Author,post_Content,post_PostTime,post_DisSM,post_DisUBB,post_DisIMG,post_AutoURL,post_AutoKEY FROM blog_Posts WHERE post_ThreadID="&ThreadID&" ORDER BY post_ID DESC"
		R_ViewThread.Open SQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		If Not(R_ViewThread.Eof And R_ViewThread.Bof) Then
			R_ViewThreadCount=R_ViewThread.RecordCount
			If R_ViewThreadCount>postPerPage Then
				R_ViewThreadContent="<div class=""content_head""><b>本主题回复较多，请 <a href=""threadview.asp?forumID="&ForumID&"&threadID="&threadID&""" target=""_blank"">点击这里</a> 查看。</div>"
			Else
				Do While Not R_ViewThread.Eof
					R_ViewThreadContent=R_ViewThreadContent&"<table class=""wordbreak""><tr><td><div class=""content_head""><span class=""smalltxt""><b><image src=""images/icon_quote.gif"" border=""0"" align=""absmiddle"">&nbsp;<a href=""member.asp?action=view&memName="&Server.URLEncode(R_ViewThread("post_Author"))&""">"&R_ViewThread("post_Author")&"</a> 于 "&DateToStr(R_ViewThread("post_PostTime"),"Y-m-d H:I A")&" 发表：</b></span></div><div class=""content_main""><span class=""smalltxt"">"&Ubbcode(HTMLEncode(R_ViewThread("post_Content")),R_ViewThread("post_DisSM"),R_ViewThread("post_DisUBB"),R_ViewThread("post_DisIMG"),R_ViewThread("post_AutoURL"),R_ViewThread("post_AutoKEY"))&"</span></div></tr></td></table>"
					R_ViewThread.MoveNext
				Loop
			End If
			R_ViewThread.Close
			Set R_ViewThread=Nothing
			Response.Write("<table width=""99%"" border=""0"" cellpadding=""0"" cellspacing=""10""><tr><td valign=""top"">"&R_ViewThreadContent&"</td></tr></table>")
		End If
	End If
End Sub

Sub ForumPostSave(Action,ForumID,ThreadID)
	Dim post_Title,post_Icon,post_Content,post_DisUBB,post_DisSM,post_DisIMG,post_AutoKEY,post_AutoURL,post_IsTop,post_IsDigest,post_IsClose,TitleFontColor
	post_Title=CheckWordFilter(Request.Form("post_Title"))
	TitleFontColor=CheckStr(Request.Form("TitleFontColor"))
	post_Content=CheckWordFilter(CheckStr(Request.Form("message")))
	post_Icon=CheckStr(Request.Form("post_Icon"))
	If Request.Form("post_IsTop")="1" Then
		post_IsTop=True
	Else
		post_IsTop=False
	End If
	If Request.Form("post_IsDigest")="1" Then
		post_IsDigest=True
	Else
		post_IsDigest=False
	End If
	If Request.Form("post_IsClose")="1" Then
		post_IsClose=True
	Else
		post_IsClose=False
	End If
	If Request.Form("post_DisSM")="1" Then
		post_DisSM=True
	Else
		post_DisSM=False
	End If
	If Request.Form("post_DisUBB")="1" Then
		post_DisUBB=True
	Else
		post_DisUBB=False
	End If
	If Request.Form("post_DisIMG")="1" Then
		post_DisIMG=True
	Else
		post_DisIMG=False
	End If
	If Request.Form("post_AutoURL")="1" Then
		post_AutoURL=True
	Else
		post_AutoURL=False
	End If
	If Request.Form("post_AutoKEY")="1" Then
		post_AutoKEY=True
	Else
		post_AutoKEY=False
	End If
	If Action="thread" Then
		If post_Title<>Empty And post_Content<>Empty Then
			Dim newThreadRS,newThreadID
			Set newThreadRS=Server.CreateObject("ADODB.RecordSet")
			SQL="SELECT * FROM blog_Threads WHERE thread_ID IS NULL"
			newThreadRS.Open SQL,Conn,1,3
			newThreadRS.AddNew
			newThreadRS("thread_ForumID")=ForumID
			newThreadRS("thread_Title")=post_Title
			newThreadRS("TitleFontColor")=TitleFontColor
			newThreadRS("thread_Icon")=post_Icon
			newThreadRS("thread_Author")=memName
			newThreadRS("thread_IsTop")=post_IsTop
			newThreadRS("thread_IsDigest")=post_IsDigest
			newThreadRS("thread_IsClose")=post_IsClose
			newThreadRS("thread_MagicFace")=CheckStr(Request.Form("post_MagicFace"))
			newThreadRS("thread_LastPost")=Now()
			newThreadRS("thread_LastPoster")=memName
			newThreadRS("thread_PostNums")=1
			newThreadRS.Update
			newThreadID=newThreadRS("thread_ID")
			newThreadRS.Close
			Set newThreadRS=Nothing
			Conn.ExeCute("INSERT INTO blog_Posts(post_ThreadID,post_ForumID,post_IsTop,post_Content,post_Author,post_PostIP,post_DisSM,post_DisUBB,post_DisIMG,post_AutoURL,post_AutoKEY) VALUES ("&newThreadID&","&ForumID&",True,'"&post_Content&"','"&memName&"','"&Guest_IP&"',"&post_DisSM&","&post_DisUBB&","&post_DisIMG&","&post_AutoURL&","&post_AutoKEY&")")
			Conn.ExeCute("UPDATE blog_Forums SET forum_ThreadNums=forum_ThreadNums+1,forum_PostNums=forum_PostNums+1 WHERE forum_ID="&ForumID&"")
			Conn.ExeCute("UPDATE blog_Member SET mem_PostThreads=mem_PostThreads+1,mem_PostPosts=mem_PostPosts+1 WHERE mem_Name='"&memName&"'")
			Conn.ExeCute("UPDATE blog_Info SET blog_PostNums=blog_PostNums+1,blog_ThreadNums=blog_ThreadNums+1")
			Response.Cookies(CookieName)("memLastpost")=Now()
			SQLQueryNums=SQLQueryNums+5
			Application.Lock
			Application(CookieName&"_blog_Forums")=""
			Application.UnLock
			Response.Write("<meta http-equiv=""refresh"" content=""3;url=threadview.asp?forumID="&ForumID&"&threadID="&newThreadID&"""><div class=""msg_head"">发表主题成功</div><div class=""msg_main""><br><br><a href=""threadview.asp?forumID="&ForumID&"&threadID="&newThreadID&""">点击返回所发主题，或者3秒后自动返回</a><br><br><a href=""forumview.asp?forumID="&ForumID&""">点击返回论坛列表</a><br><br><a href=""forumview.asp"">点击返回论坛首页</a><br><br></div>")
		Else
			Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">信息不完整，请点击返回重新填写！</a><br><br><br></div>")
		End If
	ElseIf Action="reply" Then
		If post_Content<>Empty Then
			Conn.ExeCute("INSERT INTO blog_Posts(post_ThreadID,post_ForumID,post_IsTop,post_Content,post_Author,post_PostIP,post_DisSM,post_DisUBB,post_DisIMG,post_AutoURL,post_AutoKEY) VALUES ("&ThreadID&","&ForumID&",False,'"&post_Content&"','"&memName&"','"&Guest_IP&"',"&post_DisSM&","&post_DisUBB&","&post_DisIMG&","&post_AutoURL&","&post_AutoKEY&")")
			Dim newpostRS,newpostID,newpostPage
			Set newpostRS=Conn.Execute("SELECT TOP 1 post_ID FROM blog_Posts ORDER BY post_ID DESC")
			newpostID=newpostRS("post_ID")
			Set newpostRS=Nothing
			Conn.ExeCute("UPDATE blog_Threads SET thread_LastPost='"&Now()&"',thread_LastPoster='"&memName&"',thread_PostNums=thread_PostNums+1 WHERE thread_ID="&ThreadID&"")
			Conn.ExeCute("UPDATE blog_Forums SET forum_PostNums=forum_PostNums+1 WHERE forum_ID="&ForumID&"")
			Conn.ExeCute("UPDATE blog_Member SET mem_PostPosts=mem_PostPosts+1 WHERE mem_Name='"&memName&"'")
			Conn.ExeCute("UPDATE blog_Info SET blog_PostNums=blog_PostNums+1")
			Set newpostRS=Conn.Execute("SELECT thread_PostNums FROM blog_Threads WHERE thread_ID="&threadID)
			newpostPage=newpostRS("thread_PostNums")
			If newpostPage Mod Cint(postPerPage)=0 Then
				newpostPage=Int(newpostPage/postPerPage)
			Else
				newpostPage=Int(newpostPage/postPerPage)+1
			End If
			Set newpostRS=Nothing
			Response.Cookies(CookieName)("memLastpost")=Now()
			SQLQueryNums=SQLQueryNums+7
			Response.Write("<meta http-equiv=""refresh"" content=""3;url=threadview.asp?forumID="&ForumID&"&threadID="&ThreadID&"&page="&newpostPage&"#post_"&newpostID&"""><div class=""msg_head"">发表回复成功</div><div class=""msg_main""><br><br><a href=""threadview.asp?forumID="&ForumID&"&threadID="&ThreadID&"&page="&newpostPage&"#post_"&newpostID&""">点击返回所回复主题，或者3秒后自动返回</a><br><br><a href=""forumview.asp?forumID="&ForumID&""">点击返回论坛列表</a><br><br><a href=""forumview.asp"">点击返回论坛首页</a><br><br></div>")
		Else
			Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">信息不完整，请点击返回重新填写！</a><br><br><br></div>")
		End If
	End If
End Sub

Sub SelectForum
	Response.Write("<table width=""90%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr align=""center""><td colspan=""2"" class=""msg_head"">选择论坛</td></tr><form name=""form_F"" method=""post"" action=""?action=thread""><tr bgcolor=""#FFFFFF""><td width=""19%"" align=""right""><b>论坛：</b></td><td width=""81%""><select name=""forumID"" id=""forumID""><option value="""">请选择论坛</option>")
	Dim log_ForumNums,log_ForumNumI
	log_ForumNums=Ubound(Arr_Forums,2)
	For log_ForumNumI=0 To log_ForumNums
		Response.Write("<option value="""&Arr_Forums(0,log_ForumNumI)&""">"&Arr_Forums(1,log_ForumNumI)&"</option>")
	Next
	Response.Write("</select></td></tr><tr align=""center"" bgcolor=""#FFFFFF""><td colspan=""2""><input name=""F_Submit"" type=""submit"" id=""F_Submit"" value="" 确 定 ""></td></tr></form></table>")
End Sub
%></td></tr></table>
<!--#include file="footer.asp" -->