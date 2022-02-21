<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="include/library.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr valign="top">
    <td><table width="100%"  border="0" cellpadding="0" cellspacing="0" class="wordbreak">
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td valign="top"><div class="content_head"><strong>搜索结果列表</strong></div><div class="content_main"><%
		  Call CheckPost
		  Dim sAuthor
		  sAuthor=CheckStr(Request.Querystring("Author"))
		  If sAuthor=Empty Then
		  		Response.Write("<div class=""message""><a href=""default.asp"">参数错误，点击返回主页</a></div>")
		  Else
			    Dim CurPage,Url_Add
				IF CheckStr(Request.QueryString("Page"))<>Empty Then
					Curpage=CheckStr(Request.QueryString("Page"))
					If IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
				Else
					Curpage=1
				End IF
				Url_Add="?Author="&sAuthor&"&"
				Dim aPost
				Set aPost=Server.CreateObject("ADODB.RecordSet")
				SQL="SELECT P.*,T.thread_Title,T.thread_PostNums From blog_Posts AS P,blog_Threads AS T WHERE P.post_Author='"&sAuthor&"' AND T.thread_ID=P.post_ThreadID ORDER BY post_ID DESC"
				aPost.Open SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				IF aPost.EOF AND aPost.BOF Then
					Response.Write("<div class=""message""><a href=""search.asp"">对不起，没有找到跟 "&sAuthor&" 相关的内容，点击返回主页</a></div>")
				Else
					aPost.PageSize=threadPerPage
					aPost.AbsolutePage=CurPage
					Dim Search_Nums,pAuthor,newpostPage,post_IsClose
					Search_Nums=aPost.RecordCount-1
					Dim MultiPages,PageCount
					MultiPages="<div class=""content_head""><span class=""smalltxt"">"&MultiPage(Search_Nums,threadPerPage,CurPage,Url_Add)&"</span></div>"
					Do Until aPost.EOF Or PageCount=threadPerPage
						post_IsClose=aPost("post_IsClosed")
						If post_IsClose=True Then 
							post_IsClose="&nbsp;<img src=""images/icon_postclose.gif"" align=""absmiddle"">&nbsp;"
						Else
							post_IsClose="&nbsp;"
						End If
						newpostPage=aPost("thread_PostNums")
						If newpostPage Mod Cint(postPerPage)=0 Then
							newpostPage=Int(newpostPage/postPerPage)
						Else
							newpostPage=Int(newpostPage/postPerPage)+1
						End If
						pAuthor=aPost("post_Author")
						Response.Write("<div class=""content_main"">"&post_IsClose&"<a href=""threadview.asp?forumID="&aPost("post_ForumID")&"&threadID="&aPost("post_ThreadID")&"&page="&newpostPage&"#"&aPost("post_ID")&""" target=""_blank"">"&pAuthor&" 于 "&DateToStr(aPost("post_PostTime"),"A")&" 回复主题："&HTMLEncode(aPost("thread_Title"))&"</a></div>")
						aPost.MoveNext
						PageCount=PageCount+1
					Loop
					Response.Write(MultiPages)
				End IF
				aPost.Close
				Set aPost=Nothing
			End If%></div></td>
  </tr>
</table>
    </td>
    <td width="168" bgcolor="#C8C7BD" nowrap><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	%></td>
  </tr>
</table>
<!--#include file="footer.asp" -->