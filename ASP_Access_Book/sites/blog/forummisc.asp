<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
       <tr>
    <td width="160" valign="top" bgcolor="#F7F7F4" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	%>
	</td><td>
          <%Dim msg_Title,msg_Content
	If memStatus="Admin" Or memStatus="SupAdmin" Then
		Dim Action
		Action=CheckStr(Request.QueryString("action"))
		If Action="delethread" Then
			If IsInteger(Trim(Request.QueryString("forumID")))=False Or IsInteger(Trim(Request.QueryString("threadID")))=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Dim DeleThread,DeleThreadID,DeleThreadForumID,DeleThreadPostNums,DeleThreadAuthor
				Set DeleThread=Server.CreateObject("ADODB.RecordSet")
				SQL="SELECT thread_ID,thread_ForumID,thread_Author,thread_PostNums FROM blog_Threads WHERE thread_ID="&Trim(Request.QueryString("threadID"))&" AND thread_forumID="&Trim(Request.QueryString("forumID"))&""
				DeleThread.Open SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				If DeleThread.Eof And DeleThread.Bof Then
					msg_Title="出现错误"
					msg_Content="<a href=""forumview.asp"">你所要删除的主题不存在，请点击返回论坛首页</a>"
				Else
					DeleThreadID=DeleThread("thread_ID")
					DeleThreadForumID=DeleThread("thread_ForumID")
					DeleThreadPostNums=Int(DeleThread("thread_PostNums"))
					DeleThreadAuthor=DeleThread("thread_Author")
					Conn.ExeCute("DELETE  FROM blog_Threads WHERE thread_ID="&DeleThreadID&"")
					Conn.ExeCute("UPDATE blog_Forums SET forum_ThreadNums=forum_ThreadNums-1,forum_PostNums=forum_PostNums-"&DeleThreadPostNums&" WHERE forum_ID="&DeleThreadForumID&"")
					Conn.ExeCute("UPDATE blog_Info SET blog_ThreadNums=blog_ThreadNums-1,blog_PostNums=blog_PostNums-"&DeleThreadPostNums&"")
					Conn.ExeCute("UPDATE blog_Member SET mem_PostThreads=mem_PostThreads-1 WHERE mem_Name='"&DeleThreadAuthor&"'")
					SQLQueryNums=SQLQueryNums+5
					Dim DeleThreadPost
					Set DeleThreadPost=Server.CreateObject("ADODB.RecordSet")
					SQL="SELECT post_ID,post_Author FROM blog_Posts WHERE post_ThreadID="&DeleThreadID&" AND post_ForumID="&DeleThreadForumID&""
					DeleThreadPost.Open SQL,Conn,1,3
					SQLQueryNums=SQLQueryNums+1
					Do While Not DeleThreadPost.Eof
						Conn.ExeCute("UPDATE blog_Member SET mem_PostPosts=mem_PostPosts-1 WHERE mem_Name='"&DeleThreadPost("post_Author")&"'")
						SQLQueryNums=SQLQueryNums+1
						DeleThreadPost.Delete
						DeleThreadPost.MoveNext
					Loop
					DeleThreadPost.Close
					Set DeleThreadPost=Nothing
					Application.Lock
					Application(CookieName&"_blog_Forums")=""
					Application.UnLock
					msg_Title="删除主题成功"
					msg_Content="<a href=""forumview.asp"">删除主题成功，请点击返回论坛首页</a>"
				End If
				DeleThread.Close
				Set DeleThread=Nothing
			End If
		ElseIf Action="delepost" Then
			If IsInteger(Trim(Request.QueryString("forumID")))=False Or IsInteger(Trim(Request.QueryString("threadID")))=False Or IsInteger(Trim(Request.QueryString("postID")))=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Dim DelePost,DelePostID,DelePostForumID,DelePostThreadID,DelePostAuthor
				Set DelePost = Server.CreateObject("ADODB.Recordset")
				SQL= "SELECT post_ID,post_ForumID,post_ThreadID,post_Author FROM blog_Posts WHERE post_ForumID="&Trim(Request.QueryString("forumID"))&" AND post_ThreadID="&Trim(Request.QueryString("threadID"))&" AND post_ID="&Trim(Request.QueryString("postID"))&""
				DelePost.Open SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				If DelePost.Eof And DelePost.Bof Then
					msg_Title="出现错误"
					msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
				Else
					DelePostID=DelePost("post_ID")
					DelePostForumID=DelePost("post_ForumID")
					DelePostThreadID=DelePost("post_ThreadID")
					DelePostAuthor=DelePost("post_Author")
					Conn.ExeCute("DELETE  FROM blog_Posts WHERE post_ID="&DelePostID&"")
					Dim LastPostDB
					Set LastPostDB=Conn.Execute("SELECT TOP 1 post_PostTime,post_Author FROM blog_Posts WHERE post_ThreadID="&DelePostThreadID&" ORDER BY post_ID DESC")
					Dim LastPostDBTime,LastPostDBAuthor
					LastPostDBTime=LastPostDB("post_PostTime")
					LastPostDBAuthor=LastPostDB("post_Author")
					Set LastPostDB=Nothing
					Conn.ExeCute("UPDATE blog_Threads SET thread_PostNums=thread_PostNums-1,thread_LastPost=#"&LastPostDBTime&"#,thread_LastPoster='"&LastPostDBAuthor&"' WHERE thread_ID="&DelePostThreadID&"")
					Conn.ExeCute("UPDATE blog_Forums SET forum_PostNums=forum_PostNums-1 WHERE forum_ID="&DelePostForumID&"")
					Conn.ExeCute("UPDATE blog_Member SET mem_PostPosts=mem_PostPosts-1 WHERE mem_Name='"&DelePostAuthor&"'")
					Conn.ExeCute("UPDATE blog_Info SET blog_PostNums=blog_PostNums-1")
					SQLQueryNums=SQLQueryNums+5
					msg_Title="删除回复成功"
					msg_Content="<a href=""threadview.asp?forumID="&DelePostForumID&"&threadID="&DelePostThreadID&""">删除回复成功，请点击返回原主题</a>"
				End If
				DelePost.Close
				Set DelePost = Nothing
			End If
		ElseIf Action="top" Then
			Dim top_ThreadID
			top_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(top_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsTop=1 WHERE thread_ID="&top_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="置顶操作成功"
				msg_Content="<a href=""forumview.asp"">置顶操作成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="untop" Then
			Dim untop_ThreadID
			untop_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(untop_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsTop=0 WHERE thread_ID="&untop_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="置顶解除成功"
				msg_Content="<a href=""forumview.asp"">置顶解除成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="digest" Then
			Dim digest_ThreadID
			digest_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(digest_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsDigest=True WHERE thread_ID="&digest_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="精华操作成功"
				msg_Content="<a href=""forumview.asp"">精华操作成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="undigest" Then
			Dim undigest_ThreadID
			undigest_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(undigest_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsDigest=False WHERE thread_ID="&undigest_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="精华解除成功"
				msg_Content="<a href=""forumview.asp"">精华解除成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="close" Then
			Dim close_ThreadID
			close_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(close_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsClose=True WHERE thread_ID="&close_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="锁定操作成功"
				msg_Content="<a href=""forumview.asp"">锁定操作成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="unclose" Then
			Dim unclose_ThreadID
			unclose_ThreadID=Trim(Request.QueryString("threadID"))
			If IsInteger(unclose_ThreadID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Threads SET thread_IsClose=False WHERE thread_ID="&unclose_ThreadID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="锁定解除成功"
				msg_Content="<a href=""forumview.asp"">锁定解除成功，点击返回论坛首页</a>"
			End If
		ElseIf Action="closepost" Then
			Dim close_PostID
			close_PostID=Trim(Request.QueryString("postID"))
			If IsInteger(close_PostID)=False Then
				msg_Title="出现错误"
				msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
			Else
				Conn.ExeCute("UPDATE blog_Posts SET post_IsClosed=True WHERE post_ID="&close_PostID&"")
				SQLQueryNums=SQLQueryNums+1
				msg_Title="屏蔽操作成功"
				msg_Content="<a href=""threadview.asp?forumID="&Trim(Request.QueryString("forumID"))&"&threadID="&Trim(Request.QueryString("threadID"))&""">屏蔽操作成功，点击返回主题</a>"
			End If
		Else
			msg_Title="出现错误"
			msg_Content="<a href=""forumview.asp"">参数错误，请点击返回论坛首页</a>"
		End If
	Else
		msg_Title="出现错误"
		msg_Content="<a href=""logging.asp"">没有权限，请点击登陆</a>"
	End If
	Response.Write("<div class=""msg_head"">"&msg_Title&"</div><div class=""msg_content"">"&msg_Content&"</div>")%>
          <br></td>
      </tr>
    </table>
<!--#include file="footer.asp" -->