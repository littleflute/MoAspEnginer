<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
        <tr align="center"><td>
<%
If memName<>Empty Then
	Dim edit_Action,forum_ID,thread_ID,post_ID
	edit_Action=Trim(Request.QueryString("action"))
	forum_ID=Trim(Request.QueryString("forumID"))
	thread_ID=Trim(Request.QueryString("threadID"))
	post_ID=Trim(Request.QueryString("postID"))
	If IsInteger(forum_ID)=False OR IsInteger(thread_ID)=False Then
		Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">参数无效，请点击返回上一页！</a><br><br><br></div>")
	Else
		Dim EditThreadRS
		If edit_Action="thread" Then
			Set EditThreadRS=Conn.ExeCute("SELECT thread_ID,thread_Author,thread_PostTime,thread_Title,thread_Icon,thread_IsTop,thread_IsDigest,thread_IsClose,thread_MagicFace FROM blog_Threads WHERE thread_ID="&thread_ID&" AND thread_ForumID="&forum_ID&" AND thread_Author='"&memName&"'")
				SQLQueryNums=SQLQueryNums+1
			If EditThreadRS.EOF And EditThreadRS.BOF Then
				Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">你所编辑的主题不存在，请点击返回上一页！</a><br><br><br></div>")
			Else
				If memStatus<>"SupAdmin" And DateDiff("d",EditThreadRS("thread_PostTime"),Now())>2 Then
					Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">编辑主题超过时限，请点击返回上一页！</a><br><br><br></div>")
				Else
					If Request.Form("IsPostOK")<>Empty Then
						Call CheckPost
						Call ForumEditSave(edit_Action,forum_ID,thread_ID,0)
					Else
						Call ForumEdit(edit_Action,forum_ID,thread_ID,EditThreadRS("thread_Title"),EditThreadRS("thread_Author"),EditThreadRS("thread_Icon"),EditThreadRS("thread_IsTop"),EditThreadRS("thread_IsDigest"),EditThreadRS("thread_IsClose"),EditThreadRS("thread_MagicFace"),0)
					End If
				End If
			End If
			Set EditThreadRS=Nothing
		ElseIf edit_Action="reply" Then
			If IsInteger(post_ID)=False Then
				Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">参数无效，请点击返回上一页！</a><br><br><br></div>")
			Else
				Set EditThreadRS=Conn.ExeCute("SELECT P.post_Author,T.thread_Title,P.post_PostTime FROM blog_Posts AS P,blog_Threads AS T WHERE P.post_ID="&post_ID&" AND P.post_Author='"&memName&"' AND P.post_ThreadID="&thread_ID&" AND P.post_ForumID="&forum_ID&" AND T.thread_ID=P.post_ThreadID")
				SQLQueryNums=SQLQueryNums+1
				If EditThreadRS.Eof And EditThreadRS.Bof Then
					Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">你所编辑的回复不存在，请点击返回上一页！</a><br><br><br></div>")
				Else
					If (memStatus<>"SupAdmin" Or memStatus<>"Admin") And DateDiff("d",EditThreadRS("post_PostTime"),Now())>2 Then
						Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">编辑主题超过时限，请点击返回上一页！</a><br><br><br>")
					Else
						If Request.Form("IsPostOK")<>Empty Then
							Call ForumEditSave(edit_Action,forum_ID,thread_ID,post_ID)
						Else
							Call ForumEdit(edit_Action,forum_ID,thread_ID,EditThreadRS("thread_Title"),Empty,Empty,0,False,False,"",post_ID)
						End If
					End If
				End If
				Set EditThreadRS=Nothing
			End If
		Else
			Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">参数无效，请点击返回上一页！</a><br><br><br></div>")
		End If
	End If
Else
	Response.Write("<div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""logging.asp"">游客不能编辑帖子，请先登陆</a><br><br><br></div>")
End If

Sub ForumEdit(Action,ForumID,ThreadID,ThreadTitle,ThreadAuthor,ThreadIcon,ThreadIsTop,ThreadIsDigest,ThreadIsClose,ThreadMagicFace,PostID)
	Dim TitleContent,EditPostRS,ForumContent
	Response.Write("<table width=""97%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr align=""center""><td colspan=""3"" class=""msg_head""><script language=""JavaScript"" src=""include/ubbhelp.js""></script><script language=""JavaScript"" src=""include/ubbcode.js""></script>")
	If Action="thread" Then
		Set EditPostRS=Conn.ExeCute("SELECT * FROM blog_Posts WHERE post_ThreadID="&ThreadID&" AND post_ForumID="&ForumID&" AND post_IsTop=True")
		SQLQueryNums=SQLQueryNums+1
		Response.Write("编辑主题")
		If ThreadMagicFace=Empty Then ThreadMagicFace="mf_007"
		TitleContent = "<tr bgcolor=""#FFFFFF""><td width=""130"" align=""right"" nowrap><b>标题：</b></td><td width=""100%""><input name=""edit_Title"" type=""text"" id=""edit_Title"" size=""50"" value="""&EditDeHTML(UnCheckWordFilter(ThreadTitle))&""">&nbsp;|&nbsp;转移主题到： <select name=""edit_ForumID"" id=""edit_ForumID""><option value=""0"">选择论坛</option>"
		Dim thread_MoveForumNumS,thread_MoveForumNumI
		thread_MoveForumNumS=Ubound(Arr_Forums,2)
		For thread_MoveForumNumI=0 To thread_MoveForumNumS
			TitleContent = TitleContent&"<option value='"&Arr_Forums(0,thread_MoveForumNumI)&"'>"&Arr_Forums(1,thread_MoveForumNumI)&"</option>"
		Next
		TitleContent = TitleContent&"</select></td></tr></tr><tr bgcolor=""#FFFFFF""><td align=""right"" nowrap><b>图标：</b></td><td width=""100%""><input type=""radio"" name=""edit_Icon"" value="""" "
		If ThreadIcon=Empty Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"> 无  <input type=""radio"" name=""edit_Icon"" value=""icon1.gif"" "
		If ThreadIcon="icon1.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon1.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon2.gif"" "
		If ThreadIcon="icon2.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon2.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon3.gif"" "
		If ThreadIcon="icon3.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon3.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon4.gif"" "
		If ThreadIcon="icon4.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon4.gif""> <input type=""radio"" name=""edit_Icon"" value=""icon5.gif"" "
		If ThreadIcon="icon5.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon5.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon6.gif"" "
		If ThreadIcon="icon6.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon6.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon7.gif"" "
		If ThreadIcon="icon7.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon7.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon8.gif"" "
		If ThreadIcon="icon8.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon8.gif"" align=""absmiddle""> <input type=""radio"" name=""edit_Icon"" value=""icon9.gif"" "
		If ThreadIcon="icon9.gif" Then TitleContent=TitleContent&"checked"
		TitleContent=TitleContent&"><img src=""images/threadicon/icon9.gif"" align=""absmiddle""></td></tr>"
		If memStatus="SupAdmin" Or memStatus="Admin" Then 
			TitleContent=TitleContent&"<tr bgcolor=""#FFFFFF""><td align=""right""><strong>属性：</strong></td><td colspan=""2""><input name=""edit_IsTop"" type=""checkbox"" id=""edit_IsTop"" value=""1"" "
			If ThreadIsTop=1 Then TitleContent=TitleContent&"checked"
			TitleContent=TitleContent&"> 置顶主题&nbsp;&nbsp;<input name=""edit_IsDigest"" type=""checkbox"" id=""edit_IsDigest"" value=""1"" "
			If ThreadIsDigest=True Then TitleContent=TitleContent&"checked"
			TitleContent=TitleContent&"> 精华主题&nbsp;&nbsp;<input name=""edit_IsClose"" type=""checkbox"" id=""edit_IsClose"" value=""1"" "
			If ThreadIsClose=True Then TitleContent=TitleContent&"checked"
			TitleContent=TitleContent&"> 锁定主题</td></tr>"
		End If
	ElseIf Action="reply" Then
		Set EditPostRS=Conn.ExeCute("SELECT * FROM blog_Posts WHERE post_ThreadID="&ThreadID&" AND post_ForumID="&ForumID&" AND post_ID="&PostID&"")
		SQLQueryNums=SQLQueryNums+1
		Response.Write("编辑回复")
		TitleContent="<tr bgcolor=""#FFFFFF""><td width=""112"" align=""right"" nowrap><b>主题：</b></td><td width=""100%"" colspan=""2"">"&EditDeHTML(ThreadTitle)&"</td></tr>"
	End If
	Response.Write("</td></tr>")
	If Action="thread" Then
		Response.Write("<form name=""input"" method=""post"" action=""?action="&Action&"&forumID="&ForumID&"&threadID="&ThreadID&""">")
	ElseIf Action="reply" Then
		Response.Write("<form name=""input"" method=""post"" action=""?action="&Action&"&forumID="&ForumID&"&threadID="&ThreadID&"&postID="&PostID&""">")
	End If
	Response.Write(TitleContent&"<tr bgcolor=""#FFFFFF""><td align=""right"" valign=""top"" class=""wordbreak""><b>内容：<br><br></b><div class=""siderbar_head"">表&nbsp;&nbsp;情</div><div class=""siderbar_main"" align=""left"">")
	Dim log_SmiliesListNums,log_SmiliesListNumI
	log_SmiliesListNums=Ubound(Arr_Smilies,2)
	TempVar=""
	For log_SmiliesListNumI=0 To log_SmiliesListNums
		Response.Write(TempVar&"<img style=""cursor:hand;"" onclick=""AddText('"&Arr_Smilies(2,log_SmiliesListNumI)&"');"" src=""images/smilies/"&Arr_Smilies(1,log_SmiliesListNumI)&"""/>")
		TempVar=" "
	Next
	Response.Write("</div><div style=""padding-left:15px;"" align=""left"" width=""100%""><br><br><input name=""post_DisSM"" type=""checkbox"" id=""post_DisSM"" value=""1"" ")
	If EditPostRS("post_DisSM")=True Then Response.Write("checked")
	Response.Write("> 禁止表情<br><input name=""edit_DisUBB"" type=""checkbox"" id=""edit_DisUBB"" value=""1"" ")
	If EditPostRS("post_DisUBB")=True Then Response.Write("checked")
	Response.Write("> 禁止UBB<br><input name=""edit_DisIMG"" type=""checkbox"" id=""edit_DisIMG"" value=""1"" ")
	If EditPostRS("post_DisIMG")=True Then Response.Write("checked")
	Response.Write("> 禁止图片<br><input name=""edit_AutoURL"" type=""checkbox"" id=""edit_AutoURL"" value=""1"" ")
	If EditPostRS("post_AutoURL")=True Then Response.Write("checked")
	Response.Write("> 识别链接<br><input name=""edit_AutoKEY"" type=""checkbox"" id=""edit_AutoKEY"" value=""1"" ")
	If EditPostRS("post_AutoKEY")=True Then Response.Write("checked")
	Response.Write("> 识别关键字</div></td><td colspan=""2""><table width=""98%"" border=""0"" cellspacing=""0"" cellpadding=""2""><tr><td width=""100%""><select name=""font"" onFocus=""this.selectedIndex=0"" onChange=""chfont(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>选择字体</option><option value=""宋体"">宋体</option><option value=""黑体"">黑体</option><option value=""Arial"">Arial</option><option value=""Book Antiqua"">Book Antiqua</option><option value=""Century Gothic"">Century Gothic</option><option value=""Courier New"">Courier New</option><option value=""Georgia"">Georgia</option><option value=""Impact"">Impact</option><option value=""Tahoma"">Tahoma</option><option value=""Times New Roman"">Times New Roman</option><option value=""Verdana"">Verdana</option></select><select name=""size"" onFocus=""this.selectedIndex=0"" onChange=""chsize(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>字体大小</option><option value=""-2"">-2</option><option value=""-1"">-1</option><option value=""1"">1</option><option value=""2"">2</option><option value=""3"">3</option><option value=""4"">4</option><option value=""5"">5</option><option value=""6"">6</option><option value=""7"">7</option></select><select name=""color""  onFocus=""this.selectedIndex=0"" onChange=""chcolor(this.options[this.selectedIndex].value)"" size=""1""><option value="""" selected>字体颜色</option><option value=""White"" style=""background-color:white;color:white;"">White</option><option value=""Black"" style=""background-color:black;color:black;"">Black</option><option value=""Red"" style=""background-color:red;color:red;"">Red</option><option value=""Yellow"" style=""background-color:yellow;color:yellow;"">Yellow</option><option value=""Pink"" style=""background-color:pink;color:pink;"">Pink</option><option value=""Green"" style=""background-color:green;color:green;"">Green</option><option value=""Orange"" style=""background-color:orange;color:orange;"">Orange</option><option value=""Purple"" style=""background-color:purple;color:purple;"">Purple</option><option value=""Blue"" style=""background-color:blue;color:blue;"">Blue</option><option value=""Beige"" style=""background-color:beige;color:beige;"">Beige</option><option value=""Brown"" style=""background-color:brown;color:brown;"">Brown</option><option value=""Teal"" style=""background-color:teal;color:teal;"">Teal</option><option value=""Navy"" style=""background-color:navy;color:navy;"">Navy</option><option value=""Maroon"" style=""background-color:maroon;color:maroon;"">Maroon</option><option value=""LimeGreen"" style=""background-color:limegreen;color:limegreen;"">LimeGreen</option></select></td><td rowspan=""2"" nowrap><b>模式：</b>&nbsp;<input type=""radio"" name=""mode"" value=""2"" onclick=""chmode('2')"" checked> 基本&nbsp;<input type=""radio"" name=""mode"" value=""0"" onclick=""chmode('0')""> 高级&nbsp;<input type=""radio"" name=""mode"" value=""1"" onclick=""chmode('1')""> 帮助</td></tr><tr><td><a href=""javascript:bold()""><img src=""images/ubbcode/bb_bold.gif"" border=""0"" alt=""插入粗体文本""></a> <a href=""javascript:italicize()""><img src=""images/ubbcode/bb_italicize.gif"" border=""0"" alt=""插入斜体文本""></a> <a href=""javascript:underline()""><img src=""images/ubbcode/bb_underline.gif"" border=""0"" alt=""插入下划线""></a> <a href=""javascript:center()""><img src=""images/ubbcode/bb_center.gif"" border=""0"" alt=""居中对齐""></a> <a href=""javascript:hyperlink()""><img src=""images/ubbcode/bb_url.gif"" border=""0"" alt=""插入超级链接""></a> <a href=""javascript:email()""><img src=""images/ubbcode/bb_email.gif"" border=""0"" alt=""插入邮件地址""></a> <a href=""javascript:image()""><img src=""images/ubbcode/bb_image.gif"" border=""0"" alt=""插入图像""></a> <a href=""javascript:flash()""><img src=""images/ubbcode/bb_flash.gif"" border=""0"" alt=""插入 Flash""></a> <a href=""javascript:code()""><img src=""images/ubbcode/bb_code.gif"" border=""0"" alt=""插入代码""></a> <a href=""javascript:quote()""><img src=""images/ubbcode/bb_quote.gif"" border=""0"" alt=""插入引用""></a> <a href=""javascript:list()""><img src=""images/ubbcode/bb_list.gif"" border=""0"" alt=""插入列表""></a> <a href=""javascript:wma()""><img src=""images/ubbcode/bb_wma.gif"" alt=""插入音频文件"" width=""23"" height=""22"" border=""0""></a> <a href=""javascript:wmv()""><img src=""images/ubbcode/bb_wmv.gif"" alt=""插入视频文件"" width=""23"" height=""22"" border=""0""></a></td></tr></table><table width=""98%"" border=""0"" cellpadding=""0"" cellspacing=""0""><tr valign=""top""><td><textarea name=""message"" style=""width:100%"" rows=""18"" wrap=""VIRTUAL"" id=""Message"" onSelect=""javascript: storeCaret(this);"" onClick=""javascript: storeCaret(this);"" onKeyDown=""javascript: ctlent();"" onKeyUp=""javascript: storeCaret(this);"">"&EditDeHTML(UnCheckWordFilter(EditPostRS("post_Content")))&"</textarea></td></tr></table></td></tr><tr bgcolor=""#FFFFFF""><td align=""right""><b>附件：</b></td><td colspan=""2""><iframe border=""0"" frameBorder=""0"" frameSpacing=""0"" height=""28"" marginHeight=""0"" marginWidth=""0"" noResize scrolling=""no"" width=""100%"" vspale=""0"" src=""attachment.asp""></iframe></td></tr><tr align=""center"" bgcolor=""#FFFFFF""><td colspan=""3""><input name=""IsPostOK"" type=""hidden"" value=""PostIT""><input name=""topicsubmit"" type=""submit"" value="" 提交帖子 [可按 Ctrl+Enter 发布] "" onClick=""this.disabled=true;document.input.submit();"">&nbsp;<input name=""L_Reset"" type=""reset"" id=""L_Reset"" value="" 重置帖子 ""></td></tr></form></table>")
	Set EditPostRS=Nothing
End Sub

Sub ForumEditSave(Action,ForumID,ThreadID,PostID)
	Dim edit_Title,edit_Icon,edit_Content,edit_DisUBB,edit_DisSM,edit_DisIMG,edit_AutoKEY,edit_AutoURL,edit_IsTop,edit_IsDigest,edit_IsClose,edit_Modify,edit_ForumID,MFSQL1,MFSQL2
	edit_Title=CheckWordFilter(CheckStr(Request.Form("edit_Title")))
	edit_Content=CheckWordFilter(CheckStr(Request.Form("message")))
	edit_Icon=CheckStr(Request.Form("edit_Icon"))
	edit_ForumID=Trim(Request.Form("edit_ForumID"))
	If IsInteger(edit_ForumID)=False Then edit_ForumID=0 Else edit_ForumID=Int(edit_ForumID) End If
	If edit_ForumID<>0 Then
		MFSQL1=",thread_ForumID="&edit_ForumID&""
		MFSQL2=",post_ForumID="&edit_ForumID&""
	End If
	IF Request.Form("edit_IsDigest")="1" Then
		edit_IsDigest=True
	Else
		edit_IsDigest=False
	End If
	IF Request.Form("edit_IsClose")="1" Then
		edit_IsClose=True
	Else
		edit_IsClose=False
	End If
	If Request.Form("edit_IsTop")="1" Then
		edit_IsTop=True
	Else
		edit_IsTop=False
	End If
	If Request.Form("edit_DisSM")="1" Then
		edit_DisSM=True
	Else
		edit_DisSM=False
	End If
	If Request.Form("edit_DisUBB")="1" Then
		edit_DisUBB=True
	Else
		edit_DisUBB=False
	End If
	If Request.Form("edit_DisIMG")="1" Then
		edit_DisIMG=True
	Else
		edit_DisIMG=False
	End If
	If Request.Form("edit_AutoURL")="1" Then
		edit_AutoURL=True
	Else
		edit_AutoURL=False
	End If
	If Request.Form("edit_AutoKEY")="1" Then
		edit_AutoKEY=True
	Else
		edit_AutoKEY=False
	End If
	edit_Modify="[本帖由 "&memName&" 于 "&DateToStr(Now(),"Y-m-d H:I A")&" 编辑]"
	If Action="thread" Then
		If edit_Title<>Empty And edit_Content<>Empty Then
			Conn.ExeCute("UPDATE blog_Threads SET thread_Title='"&edit_Title&"',thread_Icon='"&edit_Icon&"',thread_IsTop="&edit_IsTop&",thread_IsDigest="&edit_IsDigest&",thread_IsClose="&edit_IsClose&",thread_MagicFace='"&CheckStr(Request.Form("post_MagicFace"))&"'"&MFSQL1&" WHERE thread_ID="&ThreadID&" AND thread_ForumID="&ForumID&"")
			Conn.ExeCute("UPDATE blog_Posts SET post_Content='"&edit_Content&"',post_Modify='"&edit_Modify&"',post_DisSM="&edit_DisSM&",post_DisUBB="&edit_DisUBB&",post_DisIMG="&edit_DisIMG&",post_AutoURL="&edit_AutoURL&",post_AutoKEY="&edit_AutoKEY&""&MFSQL2&" WHERE post_ThreadID="&ThreadID&" AND post_ForumID="&ForumID&" AND post_IsTop=True")
			If edit_ForumID<>0 Then
				Dim MFPostNums
				MFPostNums=Conn.ExeCute("SELECT COUNT(post_ID) FROM blog_Posts WHERE post_ThreadID="&ThreadID&"")(0)
				Conn.ExeCute("UPDATE blog_Forums SET forum_ThreadNums=forum_ThreadNums-1,forum_PostNums=forum_PostNums-"&MFPostNums&" WHERE forum_ID="&ForumID&"")
				Conn.ExeCute("UPDATE blog_Forums SET forum_ThreadNums=forum_ThreadNums+1,forum_PostNums=forum_PostNums+"&MFPostNums&" WHERE forum_ID="&edit_ForumID&"")
				SQLQueryNums=SQLQueryNums+3
			End IF
			SQLQueryNums=SQLQueryNums+2
			Dim MFForumID
			MFForumID=ForumID
			If edit_ForumID<>0 Then MFForumID=edit_ForumID
			Response.Write("<meta http-equiv=""refresh"" content=""3;url=threadview.asp?forumID="&MFForumID&"&threadID="&ThreadID&"""><td width=""128"" align=""center"" nowrap><b><font color=""#FF0000"">编辑主题成功</font></b></td><td align=""center"" valign=""top""><div class=""msg_head"">编辑主题成功</div><div class=""msg_main""><br><br><a href=""threadview.asp?forumID="&MFForumID&"&threadID="&ThreadID&""">点击返回所编辑主题，或者3秒后自动返回</a><br><br><a href=""forumview.asp?forumID="&MFForumID&""">点击返回论坛列表</a><br><br><a href=""forumview.asp"">点击返回论坛首页</a><br><br></div></td></tr></table>")
		Else
			Response.Write("<td width=""128"" align=""center"" nowrap><b><font color=""#FF0000"">出现错误</font></b></td><td align=""center"" valign=""top""><div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">信息不完整，请点击返回重新填写！</a><br><br><br></div></td></tr></table>")
		End If
	ElseIf Action="reply" Then
		If edit_Content<>Empty Then
			Conn.ExeCute("UPDATE blog_Posts SET post_Content='"&edit_Content&"',post_Modify='"&edit_Modify&"',post_DisSM="&edit_DisSM&",post_DisUBB="&edit_DisUBB&",post_DisIMG="&edit_DisIMG&",post_AutoURL="&edit_AutoURL&",post_AutoKEY="&edit_AutoKEY&" WHERE post_ThreadID="&ThreadID&" AND post_ForumID="&ForumID&" AND post_ID="&PostID&"")
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<meta http-equiv=""refresh"" content=""3;url=threadview.asp?forumID="&ForumID&"&threadID="&ThreadID&"""><td width=""128"" align=""center"" nowrap><b><font color=""#FF0000"">编辑回复成功</font></b></td><td align=""center"" valign=""top""><div class=""msg_head"">编辑回复成功</div><div class=""msg_main""><br><br><a href=""threadview.asp?forumID="&ForumID&"&threadID="&ThreadID&""">点击返回所编辑回复的主题，或者3秒后自动返回</a><br><br><a href=""forumview.asp?forumID="&ForumID&""">点击返回论坛列表</a><br><br><a href=""forumview.asp"">点击返回论坛首页</a><br><br></div></td></tr></table>")
		Else
			Response.Write("<td width=""128"" align=""center"" nowrap><b><font color=""#FF0000"">出现错误</font></b></td><td align=""center"" valign=""top""><div class=""msg_head"">出现错误</div><div class=""msg_main""><br><br><a href=""javascript:history.go(-1);"">信息不完整，请点击返回重新填写！</a><br><br><br></div></td></tr></table>")
		End If
	End If
End Sub
%></td></tr></table>
<!--#include file="footer.asp" -->