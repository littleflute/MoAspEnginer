<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" --><table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top" nowrap bgcolor="#FFFFFF"><br>
<b>引用列表<br>
        <br>
        <%
	  Dim CurPage,Url_Add,memSQL
	  IF CheckStr(Request.QueryString("Page"))<>Empty Then
		  Curpage=CheckStr(Request.QueryString("Page"))
		  IF IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
	  Else
		  Curpage=1
	  End IF
	  memSQL="SELECT T.*,L.log_Author,L.log_Title,L.log_IsShow FROM blog_Trackback AS T,blog_Content AS L WHERE L.log_ID=T.blog_ID ORDER BY tb_PostTime DESC"
	  Url_Add="?"%>
    </b></td>
    <td width="100%" valign="top" bgcolor="#FFFFFF"><%
		Dim tb_List
		Set tb_List=Server.CreateObject("ADODB.Recordset")
		tb_List.Open memSQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		IF tb_List.EOF AND tb_List.BOF Then
			Response.Write("<table width=""98%"" border=""0"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr bgcolor=""#FFFFFF""><td>暂时没有任何引用！</td></tr></table>")
		Else
			tb_List.PageSize=8
			tb_List.AbsolutePage=CurPage
			Dim tb_Num,MultiPages,PageCount
			tb_Num=tb_List.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(tb_Num,8,CurPage,Url_Add)&"</span>"
			Response.Write(MultiPages)	
			Response.Write("<table width=""98%"" border=""0"" cellpadding=""2"" cellspacing=""0""><tr><td>")
			Do Until tb_List.EOF OR PageCount=8
				Response.Write("<fieldset align=""center"" class=""breakword"" style=""width:99%; padding:6px;""><legend style=""width:200px; background-color:#EFEFEF; padding:2px;"">&nbsp;<b>引用通告：<a href="""&tb_List("tb_URL")&""" target=""_blank"">"&tb_List("tb_Site")&"</a> 于 "&DateToStr(tb_List("tb_PostTime"),"Y-m-d H:I A")&"</b>")
				IF (memStatus="Admin" AND memName=tb_List("log_Author")) OR memStatus="SupAdmin" Then
					Response.Write("&nbsp;<a href=""trackback.asp?action=deltb&logID="&tb_List("blog_ID")&"&tbID="&tb_List("tb_ID")&""" title=""删除引用"" onClick=""winconfirm('你真的要删除这个引用吗？','trackback.asp?action=deltb&logID="&tb_List("blog_ID")&"&tbID="&tb_List("tb_ID")&"'); return false""><b><font color=""#FF0000"">×</font></b></a>")
				End IF
				Response.Write("&nbsp;</legend><img name=""HideImage"" src="""" width=""2"" height=""6"" alt="""" style=""background-color: #FFFFFF"" border=""0""><br><table class=""wordbreak""><tr><td>")
				IF tb_List("log_IsShow")=False Then
					Response.Write("<font color=""red"">隐藏日志的引用</font>")
				Else
					Response.Write("标题："&tb_List("tb_Title")&"<br>链接：<a href="""&tb_List("tb_URL")&""" target=""_blank"">"&tb_List("tb_URL")&"</a><br>摘要："&tb_List("tb_Intro"))
				End IF
				Response.Write("</td></tr></table><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF"" border=""0""><div align=""right""><a href=""blogview.asp?logID="&tb_List("blog_ID")&""">查看所引用的日志："&tb_List("log_Title")&"</a></div></fieldset><br><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF"" border=""0""><br>")
				tb_List.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("</td></tr></table>")
			Response.Write(MultiPages)	
		End IF
		tb_List.Close
		Set tb_List=Nothing
		%>
    </td>
  </tr></table>
<!--#include file="footer.asp" -->