<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" --><table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif" class="wordbreak">
  <tr>
    <td width="128" align="center" valign="top" nowrap bgcolor="#FFFFFF"><br>
<b>评论列表<br>
        <br>
        <%
	  Dim CurPage,Url_Add
	  IF CheckStr(Request.QueryString("Page"))<>Empty Then
		  Curpage=CheckStr(Request.QueryString("Page"))
		  IF IsInteger(Curpage)=False OR Curpage<0 Then Curpage=1
	  Else
		  Curpage=1
	  End IF
	  Dim commMemName,commSearch,memSQL
	  commMemName=CheckStr(Request.QueryString("memName"))
	  commSearch=CheckStr(Request.QueryString("SearchContent"))
	  IF commMemName<>Empty Then
		memSQL="SELECT C.*,L.log_Author,L.log_Title,L.log_IsShow FROM blog_Comment AS C,blog_Content AS L WHERE C.comm_Author='"&commMemName&"' AND L.log_ID=C.blog_ID ORDER BY comm_ID DESC"
	  	Response.Write(commMemName&" 所发表的评论")
	  	Url_Add="?memName="&Server.URLEncode(commMemName)&"&"
	  ElseIF commSearch<>Empty Then
	  	memSQL="SELECT C.*,L.log_Author,L.log_Title,L.log_IsShow FROM blog_Comment AS C,blog_Content AS L WHERE L.log_ID=C.blog_ID AND C.comm_Content LIKE '%"&commSearch&"%' ORDER BY comm_ID DESC"
	  	Response.Write("你所搜索的评论")
	  	Url_Add="?SearchContent="&Server.URLEncode(commSearch)&"&"
	  Else
		memSQL="SELECT C.*,L.log_Author,L.log_Title,L.log_IsShow FROM blog_Comment AS C,blog_Content AS L WHERE L.log_ID=C.blog_ID ORDER BY comm_ID DESC"
	  	Url_Add="?"
	  	Response.Write("所有评论")
	  End IF%>
    </b></td>
    <td width="100%" valign="top" bgcolor="#FFFFFF"><%
		Dim comm_List
		Set comm_List=Server.CreateObject("ADODB.Recordset")
		comm_List.Open memSQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		IF comm_List.EOF AND comm_List.BOF Then
			Response.Write("<table width=""98%"" border=""0"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr bgcolor=""#FFFFFF""><td>暂时没有任何评论！</td></tr></table>")
		Else
			comm_List.PageSize=8
			comm_List.AbsolutePage=CurPage
			Dim comm_Num,MultiPages,PageCount
			comm_Num=comm_List.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(comm_Num,8,CurPage,Url_Add)&"</span>"
			Response.Write(MultiPages)	
			Response.Write("<table width=""98%' border=""0"" cellpadding=""2"" cellspacing=""0""><tr><td>")
			Do Until comm_List.EOF OR PageCount=8
				Response.Write("<fieldset align=""center"" class=""breakword"" style=""width:99%; padding:6px;""><legend style=""width:200px; background-color:#EFEFEF; padding:2px;"">&nbsp;<b><a href=""member.asp?action=view&memName="&Server.URLEncode(comm_List("comm_Author"))&""">"&comm_List("comm_Author")&"</a> 于 "&DateToStr(comm_List("comm_PostTime"),"Y-m-d H:I A")&" 发表评论：</b>")
				IF (memStatus="Admin" AND memName=comm_List("log_Author")) OR memStatus="SupAdmin" Then
					Response.Write("&nbsp;<a href=""blogcomm.asp?action=delecomm&logID="&comm_List("blog_ID")&"&commID="&comm_List("comm_ID")&""" title=""删除评论"" onClick=""winconfirm('你真的要删除这个评论吗？','blogcomm.asp?action=delecomm&logID="&comm_List("blog_ID")&"&commID="&comm_List("comm_ID")&"'); return false""><b><font color=""#FF0000"">×</font></b></a>&nbsp;|&nbsp;<a href=""ipview.asp?ipdata="&comm_List("comm_PostIP")&""" title=""点击查看IP地址："&comm_List("comm_PostIP")&" 的来源"" target=""_blank"">IP</a>")
				End IF
				Response.Write("&nbsp;</legend><img name=""HideImage"" src="""" width=""2"" height=""6"" alt="""" style=""background-color: #FFFFFF""><br><table class=""wordbreak""><tr><td>")
				IF comm_List("log_IsShow")=False Then
					Response.Write("<font color=""red"">隐藏日志的评论</font>")
				Else
					Response.Write(UbbCode(HTMLEncode(comm_List("comm_Content")),comm_List("comm_DisSM"),comm_List("comm_DisUBB"),comm_List("comm_DisIMG"),comm_List("comm_AutoURL"),comm_List("comm_AutoKEY")))
				End IF
				Response.Write("</td></tr></table><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF""><div align=""right""><a href=""blogview.asp?logID="&comm_List("blog_ID")&""">查看所评论的日志："&comm_List("log_Title")&"</a></div></fieldset><br><img name=""HideImage"" src="""" width=""2"" height=""9"" alt="""" style=""background-color: #FFFFFF""><br>")
				comm_List.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("</td></tr></table>")
			Response.Write(MultiPages)	
		End IF
		comm_List.Close
		Set comm_List=Nothing
		%>
    </td>
  </tr></table>
<!--#include file="footer.asp" -->