<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<%Dim ForumID,Url_Add,forum_Title
	ForumID=Trim(Request.QueryString("ForumID"))
	If IsInteger(ForumID)=False Then ForumID=0
	Url_Add="?"
	Dim DigestSQL
	If Request.QueryString("ST")="D" Then
		If ForumID<>0 Then
			DigestSQL=" And thread_IsDigest=True"
		Else
			DigestSQL=" Where thread_IsDigest=True"
		End If
		Url_Add=Url_Add&"ST=D&"
	End If
	SQL="SELECT TOP 260 thread_Author,thread_LastPoster,thread_Icon,thread_ForumID,thread_ID,thread_IsTop,thread_IsDigest,thread_IsClose,thread_Title,TitleFontColor,thread_PostNums,thread_ViewNums,thread_LastPost,thread_PostTime FROM blog_Threads"&DigestSQL&" ORDER BY thread_IsTop ASC,thread_LastPost DESC"
	If ForumID<>0 Then
		SQL="SELECT  thread_Author,thread_LastPoster,thread_Icon,thread_ForumID,thread_ID,thread_IsTop,thread_IsDigest,thread_IsClose,thread_Title,TitleFontColor,thread_PostNums,thread_ViewNums,thread_LastPost,thread_PostTime FROM blog_Threads WHERE thread_ForumID="&ForumID&""&DigestSQL&" OR thread_IsTop=True ORDER BY thread_IsTop ASC,thread_LastPost DESC"
		Url_Add=Url_Add&"forumID="&ForumID&"&"
	End If%><table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="wordbreak">
  <tr>
    <td height="23"></td>
  </tr>
  <tr>
    <td valign="top">
      <%
	Response.Write("<div class=""content_head""><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td width=""100%""><a href=""forumview.asp"">51pcbook BBS</a>"&ForumTitle(ForumID)&"</td><td nowrap><b>")
	If Request.QueryString("ST")="D" Then
		If ForumID<>0 Then
			Response.Write("<a href=""forumview.asp?forumID="&ForumID&"""><img src=""images/thread_showd.gif"" align=""absmiddle"" border=""0"">&nbsp;显示所有主题</a>")
		Else
			Response.Write("<a href=""forumview.asp""><img src=""images/thread_showd.gif"" align=""absmiddle"" border=""0"">&nbsp;显示所有主题</a>")
		End If
	Else
		Response.Write("<a href=""forumview.asp"&Url_Add&"ST=D""><img src=""images/thread_showd.gif"" align=""absmiddle"" border=""0"">&nbsp;显示精华主题</a>")
	End If
	Response.Write("</b>&nbsp;</td></tr></table></div><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td align=""right""></td></tr></table>")
Dim CurPage
Call GetCurPage

Dim ThreadList
Set ThreadList=Server.CreateObject("Adodb.Recordset")
ThreadList.Open SQL,CONN,1,1
SQLQueryNums=SQLQueryNums+1
If ThreadList.EOF AND ThreadList.BOF Then 
	Response.Write("<div class=""message"">暂时没有主题</div>")
Else
	Dim Thread_Num,thread_Author,thread_LastPoster,thread_Icon,thread_forumID,thread_ID,thread_IsDigest,thread_IsTop,thread_IsClose,thread_PostNums,thread_PostPage
	ThreadList.PageSize=threadPerPage
	ThreadList.AbsolutePage=CurPage
	Thread_Num=ThreadList.RecordCount
	Dim MultiPages,PageCount
	MultiPages="<span class=""smalltxt"">"&MultiPage(Thread_Num,threadPerPage,CurPage,Url_Add)&"</span>"
	Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0""><tr><td>"&MultiPages&"</td><td align=""right"">")
	If ForumID<>0 Then
		Response.Write("<a href=""forumpost.asp?action=thread&forumID="&ForumID&""">")
	Else
		Response.Write("<a href=""forumpost.asp?action=thread"">")
	End If
	Response.Write("<img src=""images/newthread.gif"" border=""0"" align=""absmiddle""></a></td></tr></table><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table><table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr class=""msg_head"" style=""text-align:center;""><th>标题</th><th>作者</th><th nowrap>回复</th><th nowrap>查看</th><th>最后发表</th></tr>")
	Do Until ThreadList.Eof Or PageCount=threadPerPage
		thread_PostNums=ThreadList("thread_PostNums")
		thread_Author=ThreadList("thread_Author")
		thread_LastPoster=ThreadList("thread_LastPoster")
		thread_Icon=ThreadList("thread_Icon")
		thread_forumID=ThreadList("thread_ForumID")
		thread_ID=ThreadList("thread_ID")
		thread_IsTop=""
		thread_IsDigest=""
		thread_IsClose=""
		thread_PostPage=""
		If thread_PostNums>blogPerPage Then
			Dim thread_PostPageNumS,thread_PostPageNumI
			If thread_PostNums Mod Int(postPerPage)=0 Then
				thread_PostPageNumS=Int(thread_PostNums/postPerPage)
			Else
				thread_PostPageNumS=Int(thread_PostNums/postPerPage)+1
			End If
			For thread_PostPageNumI=1 To thread_PostPageNumS
				thread_PostPage=thread_PostPage&"&nbsp;<a href=""threadview.asp?forumID="&thread_forumID&"&threadID="&thread_ID&"&page="&thread_PostPageNumI&""">"&thread_PostPageNumI&"</a>"
			Next
		End If
		If thread_PostPage<>Empty Then thread_PostPage="<span class=""smalltxt"">&nbsp;&nbsp;(&nbsp;<img src=""images/icon_pages.gif"" align=""absmiddle"" boader=""0"">"&thread_PostPage&"&nbsp;)</span>"
		If ThreadList("thread_IsTop")=True Then thread_IsTop="&nbsp;<img src=""images/thread_top.gif"" align=""absmiddle"" border=""0"">"
		If ThreadList("thread_IsDigest")=True Then thread_IsDigest="&nbsp;<img src=""images/thread_digest.gif"" align=""absmiddle"" border=""0"">"
		If ThreadList("thread_IsClose")=True Then thread_IsClose="&nbsp;<img src=""images/thread_close.gif"" align=""absmiddle"" border=""0"">"
		Response.Write("<tr bgcolor=""#FFFFFF""><td width=""100%"" style=""font-size:12px;"">")
		If Trim(thread_Icon)=Empty Then thread_Icon="../newwin.gif"
		Response.Write("<a href=""threadview.asp?forumID="&thread_forumID&"&threadID="&thread_ID&""" target=""_blank""><img src=""images/threadicon/"&thread_Icon&""" border=""0"" align=""absmiddle"" alt=""点击在新窗口打开""></a>&nbsp;<a href=""threadview.asp?forumID="&thread_forumID&"&threadID="&thread_ID&"""><font color='"&ThreadList("TitleFontColor")&"'>"&HTMLEncode(cutStr(ThreadList("thread_Title"),60))&"</font></a>"&thread_IsTop&thread_IsDigest&thread_IsClose&thread_PostPage&"</td><td nowrap class=""ssmalltxt"" align=""center""><a href=""member.asp?action=view&memName="&Server.URLEncode(thread_Author)&""">"&thread_Author&"</a><br>"&DateToStr(ThreadList("thread_PostTime"),"Y-m-d")&"</td><td nowrap class=""ssmalltxt"" align=""center"">"&thread_PostNums&"</td><td nowrap class=""smalltxt"" align=""center"">"&ThreadList("thread_ViewNums")&"</td><td nowrap class=""ssmalltxt"" align=""right"">"&DateToStr(ThreadList("thread_LastPost"),"Y-m-d H:I A")&"<br>by&nbsp;<a href=""member.asp?action=view&memName="&Server.URLEncode(thread_LastPoster)&""">"&thread_LastPoster&"</a></td></tr>")
		ThreadList.MoveNext
		PageCount=PageCount+1
	Loop
	Response.Write("</table><table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" height=""6""><tr><td></td></tr></table>")
End If
ThreadList.Close
Set ThreadList=Nothing
Response.Write(MultiPages)
%></td>
  </tr>
</table></td>
    <td width="160" valign="top" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Call ForumList(ForumID)
	Call T_Forum
	Call R_Forum
	Call SiteInfo
	Call forumSearch
	%>
	
	</td>
  </tr>
</table>
<!--#include file="footer.asp" -->