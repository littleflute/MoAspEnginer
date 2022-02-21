

<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="header.asp" -->
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="160" valign="top" nowrap bgcolor="#C8C7BD"><%
	Call MemberCenter
	Response.Write("<br>")
	Call SiteInfo
	Response.Write("<br>")
	Call blogSearch%>    
    <br></td>
    <td width="100%" valign="top" bgcolor="#FFFFFF" class="wordbreak"><%Dim blog_Links,blog_ImgLinks,blog_TxtLinks
	Set blog_Links=Conn.ExeCute("SELECT * FROM blog_Links WHERE link_IsShow=True AND link_IsMain=False ORDER BY link_Order ASC")
	SQLQueryNums=SQLQueryNums+1
	IF Not(blog_Links.EOF And blog_Links.BOF) Then
		Do While Not blog_Links.EOF
			IF blog_Links("link_Image")<>Empty Then
				blog_ImgLinks=blog_ImgLinks&"<span style=""width:95px;""><a href="""&blog_Links("link_URL")&""" target=""_blank""><img src="""&blog_Links("link_Image")&""" alt="""&blog_Links("link_Name")&""" border=""0"" align=""absmiddle""></a></span>"
			Else
				blog_TxtLinks=blog_TxtLinks&"<span class=""hyperlink""><a href="""&blog_Links("link_URL")&""" target=""_blank"">"&blog_Links("link_Name")&"</a></span>"
			End IF
			blog_Links.MoveNext
		Loop
	End IF
	Set blog_Links=Nothing%><div class="content_head"><strong>图片链接：</strong></div><div class="content_main"><%=blog_ImgLinks%><br /><br /></div>
      <div class="content_head"><strong>文本链接：</strong></div><div class="content_main">
      <%=blog_TxtLinks%><br /><br /></div>
    <div class="content_head"><strong>友情链接加入说明：</strong></div><div class="content_main">
    如果你想跟本BLOG建立友情链接，请<strong>先在你的站点上加上本BLOG的友情链接</strong>，然后在这里登陆你的站点！谢谢合作！</div>
    <%If CheckStr(Request.Form("link_Name"))=Empty OR CheckStr(Request.Form("link_URL"))=Empty Then%><table width="95%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC"><form name="blogLinksAdd" method="post" action="">
      <tr bgcolor="#FFFFFF">
        <td width="8%" nowrap>名称：</td>
        <td width="36%"><input name="link_Name" type="text" id="link_Name" size="25"> <strong>必填</strong></td>
        <td width="7%" nowrap>地址：</td>
        <td width="49%"><input name="link_URL" type="text" id="link_URL" size="38"> <strong>必填</strong></td>
        </tr>
      <tr bgcolor="#EFEFEF">
        <td nowrap bgcolor="#FFFFFF">图片：
        </td>
        <td colspan="2" bgcolor="#FFFFFF"><input name="link_IMG" type="text" id="link_IMG" size="42"></td>
        <td bgcolor="#FFFFFF"><input type="submit" name="Submit" value=" 提交你的站点 "></td>
        </tr></form>
    </table>
    <%Else
	Dim link_LastOrder,link_LastQuery
	Set link_LastQuery=Conn.ExeCute("SELECT TOP 1 link_Order FROM blog_Links ORDER BY link_Order DESC")
	IF link_LastQuery.EOF And link_LastQuery.BOF Then
		link_LastOrder=1
	Else
		link_LastOrder=link_LastQuery(0)+1
	End IF
	Set link_LastQuery=Nothing
	Conn.ExeCute("INSERT INTO blog_Links(link_Name,link_URL,link_Image,link_Order) VALUES ('"&CheckStr(Request.Form("link_Name"))&"','"&CheckStr(Request.Form("link_URL"))&"','"&CheckStr(Request.Form("link_IMG"))&"',"&link_LastOrder&")")
	SQLQueryNums=SQLQueryNums+2%>
    <table width="95%" border="0" cellpadding="22" cellspacing="1" bgcolor="#CCCCCC">
      <tr bgcolor="#EFEFEF">
        <td colspan="4" align="center" bgcolor="#FFFFFF"><strong>你的站点已经提交成功，请等待我的验证！ </strong></td>
      </tr>
    </table>
	<%End IF%>
    <br>
</td>
  </tr>
</table>
<!--#include file="footer.asp" -->