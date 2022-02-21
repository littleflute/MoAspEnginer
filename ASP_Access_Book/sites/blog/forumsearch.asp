<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="768" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr><td width="160" nowrap class="siderbar"><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	%></td>
    <td valign="top"><div class="content_head"><strong>搜索结果列表</strong></div><div class="content_main"><%
		  Call CheckPost
		  Dim forum_SearchContent,Is_Post
		  forum_SearchContent=CheckStr(Request.Form("SearchContent"))
		  Is_Post=CheckStr(Request.Form("Is_Post"))
		  If forum_SearchContent=Empty Then forum_SearchContent=CheckStr(Request.Querystring("SearchContent"))
		  If Is_Post=Empty Then Is_Post=CheckStr(Request.Querystring("IsPost"))
		  If forum_SearchContent=Empty Then
		  		Response.Write("<div class=""message""><a href=""forumview.asp"">对不起，请输入搜索关键字，点击返回搜索页面</a></div>")
		  Else
			    Dim CurPage,Url_Add,SQLFiltrate
				IF CheckStr(Request.QueryString("Page"))<>Empty Then
					Curpage=CheckStr(Request.QueryString("Page"))
					IF IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
				Else
					Curpage=1
				End IF
				Url_Add = "?SearchContent="&Server.URLEncode(forum_SearchContent)&"&"
				SQL="SELECT * From blog_Threads WHERE thread_ID=0 ORDER BY thread_LastPost DESC"
				If Is_Post = "1" Then
					Url_Add = Url_Add&"IsPost=1&"
					SQL = "SELECT T.* FROM blog_Threads AS T,blog_Posts AS P WHERE (P.post_Content LIKE '%"&forum_SearchContent&"%' OR T.thread_Title LIKE '%"&forum_SearchContent&"%') AND T.thread_ID=P.post_ThreadID ORDER BY T.thread_LastPost DESC"
				Else
					SQL="SELECT * FROM blog_Threads WHERE thread_Title LIKE '%"&forum_SearchContent&"%' ORDER BY thread_LastPost DESC"
				End If
				Dim forum_Search
				Set forum_Search=Server.CreateObject("ADODB.RecordSet")
				forum_Search.Open SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				IF forum_Search.EOF AND forum_Search.BOF Then
					Response.Write("<div class=""message""><a href=""forumview.asp"">对不起，没有找到跟 "&forum_SearchContent&" 相关的内容，点击搜索页面</a></div>")
				Else
					forum_Search.PageSize=20
					forum_Search.AbsolutePage=CurPage
					Dim Search_Nums,thread_Author
					Search_Nums=forum_Search.RecordCount-1
					Dim MultiPages,PageCount
					MultiPages="<div class=""content_head""><span class=""smalltxt"">"&MultiPage(Search_Nums,20,CurPage,Url_Add)&"</span></div>"
					Do Until forum_Search.EOF Or PageCount=20
						thread_Author=forum_Search("thread_Author")
						Response.Write("<div class=""content_main"">&nbsp;<a href=""threadview.asp?forumID="&forum_Search("thread_ForumID")&"&threadID="&forum_Search("thread_ID")&""" target=""_blank"">"&forum_Search("thread_Title")&"</a></div>")
						forum_Search.MoveNext
						PageCount=PageCount+1
					Loop
					Response.Write(MultiPages)
				End IF
				forum_Search.Close
				Set forum_Search=Nothing
			End If%></div></td>
  </tr>
</table>
<!--#include file="footer.asp" -->