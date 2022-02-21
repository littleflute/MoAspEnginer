<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="header.asp" --><table width="776" border="0" cellpadding="9" cellspacing="0" align="center" bgcolor="#FFFFFF">
  <tr>
    <td width="138" align="center" valign="top" nowrap class="rtd"><br>
<br>
<b><font color="#FF0000">BLOG 搜索</font></b><br>
<br></td>
    <td width="100%" valign="top"><div class="content_head"><strong>搜索结果列表</strong></div><div class="content_main"><%
		  Dim blog_SearchContent,Is_Title,Is_Content
		  blog_SearchContent=CheckStr(Request.Form("SearchContent"))
		  Is_Title=CheckStr(Request.Form("Is_Title"))
		  Is_Content=CheckStr(Request.Form("Is_Content"))
		  IF blog_SearchContent=Empty Then blog_SearchContent=CheckStr(Request.Querystring("SearchContent"))
		  IF Is_Title=Empty Then Is_Title=CheckStr(Request.Querystring("IsTitle"))
		  IF Is_Content=Empty Then Is_Content=CheckStr(Request.Querystring("IsContent"))
		  IF blog_SearchContent=Empty Then
		  		Response.Write("<div class=""message""><a href=""default.asp"">对不起，请输入搜索关键字，点击返回主页</a></div>")
		  Else
			    Dim CurPage,Url_Add,SQLFiltrate
				IF CheckStr(Request.QueryString("Page"))<>Empty Then
					Curpage=CheckStr(Request.QueryString("Page"))
					IF IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
				Else
					Curpage=1
				End IF
				Url_Add = "?SearchContent="&Server.URLEncode(blog_SearchContent)&"&"
				If Is_Title = "1" And Is_Content = "1" Then
					Url_Add = Url_Add&"IsTitle=1&IsContent=1&"
					SQLFiltrate = "WHERE log_Title LIKE '%"&blog_SearchContent&"%' OR log_Content LIKE '%"&blog_SearchContent&"%'"
				ElseIf	Is_Title = "1" Then
					Url_Add = Url_Add&"IsTitle=1&"
					SQLFiltrate = "WHERE log_Title LIKE '%"&blog_SearchContent&"%'"
				ElseIf Is_Content = "1" Then 
					Url_Add = Url_Add&"Is_Content=1&"
					SQLFiltrate = "WHERE log_Content LIKE '%"&blog_SearchContent&"%'"
				End If
				Dim blog_Search
				Set blog_Search=Server.CreateObject("ADODB.RecordSet")
				SQL="SELECT * FROM blog_Content "&SQLFiltrate&" ORDER BY log_PostTime DESC"
				blog_Search.Open SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				IF blog_Search.EOF AND blog_Search.BOF Then
					Response.Write("<div class=""message""><a href=""default.asp"">对不起，没有找到跟 "&HTMLEncode(blog_SearchContent)&" 相关的内容，点击返回主页</a></div>")
				Else
					blog_Search.PageSize=20
					blog_Search.AbsolutePage=CurPage
					Dim Search_Nums
					Search_Nums=blog_Search.RecordCount-1
					Dim MultiPages,PageCount
					MultiPages="<div class=""content_head""><span class=""smalltxt"">"&MultiPage(Search_Nums,20,CurPage,Url_Add)&"</span></div>"
					Do Until blog_Search.EOF OR PageCount=20
						Response.Write("<div class=""object_main"">&nbsp;<a href=""blogview.asp?logID="&blog_Search("log_ID")&""" target=""_blank"">"&blog_Search("log_Title")&"</a></div>")
						blog_Search.MoveNext
						PageCount=PageCount+1
					Loop
					Response.Write(MultiPages)
				End IF
				blog_Search.Close
				Set blog_Search=Nothing
			End IF%></td>
        </tr>
      </table></td>
  </tr>
</table>
<!--#include file="footer.asp" -->