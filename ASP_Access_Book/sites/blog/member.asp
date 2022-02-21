<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="include/ubbcode.asp" -->
<!--#include file="header.asp" -->
<%IF Request.QueryString("action")="edit" AND memName<>Empty Then%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center"><strong>用户编辑</strong></td>
    <td valign="top" align="center"><table width="95%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
      <tr bgcolor="#EFEFEF">
        <td colspan="2"><b>编辑个人资料</b></td>
        </tr><%
IF Request.QueryString("edit")="post" Then
	Response.Write("<tr bgcolor=""#FFFFFF""><td colspan=""2"">")
	IF CheckStr(Request.Form("mem_OldPassword"))=Empty Then
		Response.Write("<br><br><center><b>对不起，旧密码必须填写<br><br><a href='javascript:history.go(-1);'>点击返回上一页</a></b></center><br><br>")
	Else
		Dim mem_Edit
		Set mem_Edit=Conn.ExeCute("SELECT * FROM blog_Member WHERE mem_Name='"&memName&"' AND mem_PassWord='"&md5(CheckStr(Request.Form("mem_OldPassword")))&"'")
		SQLQueryNums=SQLQueryNums+1
		IF mem_Edit.EOF AND mem_Edit.BOF Then
			Response.Write("<br><br><center><b>对不起，旧密码不符<br><br><a href='javascript:history.go(-1);'>点击返回上一页</a></b></center><br><br>")
		Else
			IF CheckStr(Request.Form("mem_NewPassword"))<>CheckStr(Request.Form("mem_RePassword"))  Then
				Response.Write("<br><br><center><b>对不起，新密码与确认密码不符<br><br><a href='javascript:history.go(-1);'>点击返回上一页</a></b></center><br><br>")
			Else
				Dim SQL_Add,mem_HideEmail
				IF CheckStr(Request.Form("mem_NewPassword"))<>Empty Then SQL_Add=",mem_Password='"&md5(CheckStr(Request.Form("mem_NewPassword")))&"'"
				mem_HideEmail=CheckStr(Request.Form("mem_HideEmail"))
				IF mem_HideEmail="1" Then
					mem_HideEmail=True
				Else
					mem_HideEmail=False
				End IF
				Conn.ExeCute("UPDATE blog_Member SET mem_Sex="&CheckStr(Request.Form("mem_Sex"))&",mem_Email='"&CheckStr(Request.Form("mem_Email"))&"',mem_HideEmail="&mem_HideEmail&",mem_QQ='"&CheckStr(Request.Form("mem_QQ"))&"',mem_HomePage='"&CheckStr(Request.Form("mem_HomePage"))&"',mem_Intro='"&CheckStr(Request.Form("mem_Intro"))&"'"&SQL_Add&" WHERE mem_Name='"&memName&"'")
				SQLQueryNums=SQLQueryNums+1
				Response.Write("<br><br><center><b>更新资料成功<br><br><a href='member.asp?action=view&memName=&Server.URLEncode(memname)&""'>点击查看新资料</a></b></center><br><br>")
			End IF
		End IF
		Set mem_Edit=Nothing
	End IF
	Response.Write("</td></tr>")
Else
	 Dim mem_EditView
	 Set mem_EditView=Conn.ExeCute("SELECT * FROM blog_Member WHERE mem_Name='"&memName&"'")
	 SQLQueryNums=SQLQueryNums+1%>
	 <tr bgcolor="#FFFFFF"><form action="member.asp?action=edit&edit=post" method="post" name="Submit" id="Submit">
        <td width="98" align="right" nowrap>名称：</td>
        <td width="100%">&nbsp;<%=memName%></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>旧密码：</td>
        <td>&nbsp;<input name="mem_OldPassword" type="password" id="mem_OldPassword" size="16">
          <span class="htd"><b><font color="#FF0000">&nbsp;*</font></b> 修改资料必须输入原密码</span></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>新密码：</td>
        <td>&nbsp;<input name="mem_NewPassword" type="password" id="mem_NewPassword" size="16">
          &nbsp;<span class="htd"><b><font color="#FF0000">*</font></b> 密码必须是6-16位</span></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>确认密码：</td>
        <td>&nbsp;<input name="mem_RePassword" type="password" id="mem_RePassword" size="16"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>性别：</td>
        <td><input name="mem_Sex" type="radio" value="1" <%IF mem_EditView("mem_Sex")=1 Then Response.Write("checked")%>>
          男
            <input type="radio" name="mem_Sex" value="2" <%IF mem_EditView("mem_Sex")=2 Then Response.Write("checked")%>>
            女
            <input type="radio" name="mem_Sex" value="0" <%IF mem_EditView("mem_Sex")=0 Then Response.Write("checked")%>>
            保密</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>邮箱：</td>
        <td>&nbsp;<input name="mem_Email" type="text" id="mem_Email" value="<%=mem_EditView("mem_Email")%>" size="30">
          &nbsp;
          <input name="mem_HideEmail" type="checkbox" id="mem_HideEmail" value="1" <%IF mem_EditView("mem_HideEmail")=True Then Response.Write("Checked")%>>
          隐藏邮箱</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>腾讯QQ：</td>
        <td>&nbsp;<input name="mem_QQ" type="text" id="mem_QQ" value="<%=mem_EditView("mem_QQ")%>"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>个人主页：</td>
        <td>&nbsp;<input name="mem_HomePage" type="text" id="mem_HomePage" value="<%=mem_EditView("mem_HomePage")%>" size="35"></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>个人简介：</td>
        <td>&nbsp;<textarea name="mem_Intro" cols="50" rows="5" wrap="VIRTUAL" id="mem_Intro"><%=mem_EditView("mem_Intro")%></textarea></td>
      </tr>
      <tr align="center" bgcolor="#FFFFFF">
        <td colspan="2">
          <input type="submit" name="Submit" value=" 确定编辑 ">
        </td>
        </tr></form><%Set mem_EditView=Nothing
		End IF%>
    </table></td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="view" AND CheckStr(Request.QueryString("memName"))<>Empty Then%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center"><strong>用户查看</strong></td>
    <td valign="top" align="center"><%Dim mem_View
	Set mem_View=Conn.ExeCute("SELECT * FROM blog_Member WHERE mem_Name='"&CheckStr(Request.QueryString("memName"))&"'")
	SQLQueryNums=SQLQueryNums+1
	IF mem_View.EOF AND mem_View.BOF Then
		Response.Write("<br><br><center><b>对不起，你所查询的用户不存在<br><br><a href='javascript:history.go(-1);'>点击返回上一页</a></b></center><br><br>")
	Else%><table width="95%" border="0" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
      <tr bgcolor="#EFEFEF">
        <td colspan="2"><b>用户 <font color="#FF0000"><%=mem_View("mem_Name")%></font> 的详细资料：</b></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">名称：</td>
        <td>&nbsp;<%=mem_View("mem_Name")%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td width="98" align="right" nowrap>性别：</td>
        <td width="100%">&nbsp;<%IF mem_View("mem_Sex")=1 Then
			Response.Write("<img src='images/male.gif' border='0' align='absmiddle'>&nbsp;帅哥")
		ElseIF mem_View("mem_Sex")=2 Then
			Response.Write("<img src='images/female.gif' border='0' align='absmiddle'>&nbsp;美女")
		Else
			Response.Write("未知")
		End IF%></td>
        </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">邮箱：</td>
        <td>&nbsp;<%IF mem_View("mem_HideEmail")=False Then
				Response.Write("<a href='mailto:"&HTMLEncode(mem_View("mem_Email"))&"' alt='点击发送 Email'>"&HTMLEncode(mem_View("mem_Email"))&"</a>")
			Else
				Response.Write("此用户信箱已隐藏")
			End IF%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">腾讯QQ：</td>
        <td>&nbsp;<%="<a href='http://friend.qq.com/cgi-bin/friend/user_show_info?ln="&HTMLEncode(mem_View("mem_QQ"))&"' target='_blank' alt='点击查看QQ信息'>"&HTMLEncode(mem_View("mem_QQ"))&"</a>"%>&nbsp;</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right" nowrap>个人主页：</td>
        <td>&nbsp;<%="<a href='"&HTMLEncode(mem_View("mem_HomePage"))&"' target='_blank' alt='点击浏览主页'>"&HTMLEncode(mem_View("mem_HomePage"))&"</a>"%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">发表日志：</td>
        <td>&nbsp;<%=mem_View("mem_PostLogs")%> 篇</td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">发表评论：</td>
        <td>&nbsp;<%=mem_View("mem_PostComms")%> 篇&nbsp;&nbsp;&nbsp;&nbsp;[<a href="commlist.asp?memName=<%=Server.URLEncode(mem_View("mem_Name"))%>">点击浏览 <%=mem_View("mem_Name")%> 发表的所有评论</a>]</td>
      </tr>
	  <tr bgcolor="#FFFFFF">
          <td align="right">发表主题：</td>
          <td>&nbsp;<%=mem_View("mem_PostThreads")%> 篇</td>
        </tr>
        <tr bgcolor="#FFFFFF">
          <td align="right">发表帖子：</td>
          <td>&nbsp;<%=mem_View("mem_PostPosts")%> 篇&nbsp;&nbsp;&nbsp;&nbsp;[<a href="postlist.asp?Author=<%=Server.URLEncode(mem_View("mem_Name"))%>">点击浏览 <%=mem_View("mem_Name")%> 发表的所有帖子</a>]</td>
        </tr>
	  <tr bgcolor="#FFFFFF">
        <td align="right">发表留言：</td>
        <td>&nbsp;<%=mem_View("mem_PostGBNums")%> 篇&nbsp;&nbsp;&nbsp;&nbsp;[<a href="guestbook.asp?memName=<%=Server.URLEncode(mem_View("mem_Name"))%>">点击浏览 <%=mem_View("mem_Name")%> 发表的所有留言</a>]</td>
	</tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">个人介绍：</td>
        <td><%=UBBCode(HTMLEncode(mem_View("mem_Intro")),1,0,1,0,0)%></td>
      </tr>
      <tr bgcolor="#FFFFFF">
        <td align="right">注册时间：</td>
        <td>&nbsp;<%=DateToStr(mem_View("mem_RegTime"),"Y-m-d H:I A")%></td>
        </tr>
    </table>
    <%End IF
Set mem_View=Nothing%></td>
  </tr>
</table>
<%Else%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top"><strong><br>
      <br>
    用户列表</strong></td>
    <td valign="top" align="center"><table width="100%" border="0" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
      <tr bgcolor="#EFEFEF">
        <td width="10%" align="center" nowrap><strong>编号</strong></td>
        <td width="22%" nowrap><strong>名称</strong></td>
        <td width="7%" align="center"><strong>邮箱</strong></td>
        <td width="7%" align="center"><strong>QQ</strong></td>
        <td width="7%" align="center"><strong>主页</strong></td>
        <td width="10%" align="center"><strong>日志</strong></td>
        <td width="10%" align="center"><strong>评论</strong></td>
		<td width="10%" align="center"><strong>留言</strong></td>
		<td nowrap><strong>注册时间</strong></td>
      </tr><%Dim CurPage
	  IF CheckStr(Request.QueryString("Page"))<>Empty Then
			Curpage=Cint(CheckStr(Request.QueryString("Page")))
			IF Curpage<0 Then Curpage=1
	  Else
			Curpage=1
	  End IF
	  Dim blog_Member
	  Set blog_Member=Server.CreateObject("Adodb.Recordset")
	  SQL="SELECT * FROM blog_Member ORDER BY mem_RegTime DESC"
	  blog_Member.Open SQL,CONN,1,1
	  SQLQueryNums=SQLQueryNums+1
	  IF blog_Member.EOF AND blog_Member.BOF Then
	  		Response.Write("<tr bgcolor=""#FFFFFF""><td colspan=""8"" align=""center"" height=""48"">暂时没有用户</td></tr>")
	  Else
	  		Dim list_MemName
	  		blog_Member.Pagesize=15
			blog_Member.AbsolutePage=CurPage
			Dim Member_Nums,PageCount,MultiPages
			Member_Nums=blog_Member.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(Member_Nums,15,CurPage,"?")&"</span>"
	  		Do Until blog_Member.EOF OR PageCount=15
				list_MemName=blog_Member("mem_Name")
				Response.Write("<tr bgcolor=""#FFFFFF""><td>"&blog_Member("mem_ID")&"</td><td nowrap><a href=""member.asp?action=view&memName="&Server.URLEncode(list_MemName)&""" alt=""点击查看会员 "&list_MemName&" 的详细资料"">"&list_MemName&"</a></td><td align=""center"">")
				If blog_Member("mem_HideEmail")=False And blog_Member("mem_Email")<>Empty Then Response.Write("<a href=""mailto:"&HtmlEncode(blog_Member("mem_Email"))&"""><img src=""images/icon_email.gif"" border=""0"" align=""absmiddle"" alt=""点击给会员 "&list_MemName&" 发送 Email""></a>")
				Response.Write("</td><td nowrap align=""center"">")
				If blog_Member("mem_QQ")<>Empty Then Response.Write("<a href=""http://friend.qq.com/cgi-bin/friend/user_show_info?ln="&HtmlEncode(blog_Member("mem_QQ"))&""" target=""_blank""><img src=""images/icon_qq.gif"" border=""0"" align=""absmiddle"" alt=""点击查看会员 "&list_MemName&" 的QQ资料""></a>")
				Response.Write("</td><td align=""center"">")
				If blog_Member("mem_Homepage")<>Empty Then Response.Write("<a href='"&HTMLEncode(blog_Member("mem_Homepage"))&"' target=""_blank""><img src=""images/icon_home.gif"" border=""0"" align=""absmiddle"" alt=""点击访问会员 "&list_MemName&" 的个人主页""></a>")
				Response.Write("</td><td>"&blog_Member("mem_PostLogs")&"</td><td nowarp>"&blog_Member("mem_PostComms")&"</td><td nowarp>"&blog_Member("mem_PostGBNums")&"</td><td width=""100%"" nowrap>"&DateToStr(blog_Member("mem_RegTime"),"Y-m-d H:I A")&"</td></tr>")
				blog_Member.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("<tr bgcolor=""#FFFFFF""><td colspan=""9"">"&MultiPages&"</td></tr>")
		End IF
		blog_Member.Close
		Set blog_Member=Nothing%>
    </table></td>
  </tr>
</table>
<%End IF%>
<!--#include file="footer.asp" -->