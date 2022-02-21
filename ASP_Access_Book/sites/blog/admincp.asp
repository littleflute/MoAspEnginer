<!--#include file="commond.asp" -->
<!--#include file="include/function.asp" -->
<!--#include file="include/md5code.asp" -->
<!--#include file="header.asp" -->
<%On Error Resume Next%>
<table width="776" border="0" align="center" cellpadding="4" cellspacing="6" background="images/blog_main.gif">
  <tr>
    <td width="128" align="center" valign="top" nowrap align="center" bgcolor="#FFFFFF"><br>
    <br><div class="msg_head">管理面板导航</div><div class="msg_main"><a href="admincp.asp"><b>管理首页</b></a><br />
<%IF Session("Admin")<>Empty Then%>
<a href="admincp.asp?action=category"><b>分类管理</b></a><br />
<a href="admincp.asp?action=photocate"><b>相册分类</b></a><br>
<a href="admincp.asp?action=member"><b>会员管理</b></a><br />
<a href="admincp.asp?action=attachment"><b>附件管理</b></a><br />
<a href="admincp.asp?action=smilies"><b>表情管理</b></a><br />
<a href="admincp.asp?action=keywords"><b>关键字管理</b></a><br />
<a href="admincp.asp?action=wordfilter"><b>语言过滤</b></a><br />
<a href="admincp.asp?action=links"><b>链接管理</b></a><br />
<a href="admincp.asp?action=linkscheck"><b>链接验证</b></a><br />
<a href="admincp.asp?action=favorite"><b>书签管理</b></a><br />
<a href="admincp.asp?action=download"><b>下载管理</b></a><br>
<a href="admincp.asp?action=logout"><b>退出登录</b></a></div>
<%End IF%></div><br><br></td>
    <td width="100%" valign="top" bgcolor="#FFFFFF" align="center"><%IF memStatus<>"SupAdmin" Then%><br /><br /><br /><div class="msg_content">对不起，你没有权限进入系统管理面板<br /><br /><a href="default.asp" target="_top">点击返回首页</a></div><br /><br /><%Else%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top" align="center"><%IF Session("Admin")=Empty Then%>
      <br><br><br><br><table width="40%" border="0" align="center" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head">请输入管理员密码：</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" align="center"><br><form name="adminlogin" method="post" action="admincp.asp?action=login"><input type="password" id="adminpassword" name="adminpassword">&nbsp;<input type="submit" id="submit" name="submit" value=" 确定登陆 "></form></td>
  </tr>
</table>
<%IF Request.QueryString("action")="login" Then
	Dim AdminLogin,AdminLogin_OK
	Set AdminLogin=Conn.ExeCute("SELECT mem_PassWord,mem_Name FROM blog_Member WHERE mem_Name='"&memName&"' AND mem_PassWord='"&md5(CheckStr(Request.Form("adminpassword")))&"'")
	SQLQueryNums=SQLQueryNums+1
	IF AdminLogin.EOF AND AdminLogin.BOF Then
		AdminLogin_OK=0
	Else
		AdminLogin_Ok=1
	End IF
	Set AdminLogin=Nothing
	IF AdminLogin_Ok=1 Then Session("Admin")=memName
	Response.Redirect("admincp.asp")
End IF
Else
IF Request.QueryString("action")="logout" Then
	Session("Admin")=""
	Response.ReDirect("default.asp")
ElseIF Request.QueryString("action")="database" Then%>
<br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 数据管理</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">数据库文件路径：<%=Request.ServerVariables("APPL_PHYSICAL_PATH")&AccessPath&"\"&AccessFile%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">数据库文件空间占用：<%=GetTotalSize(Server.Mappath(AccessPath&"/"&AccessFile),"File")%></td>
  </tr>
   <tr><form action="admincp.asp?action=database&type=sqlquery" method="post">
    <td bgcolor="#FFFFFF">SQL 查询执行(一次执行一个查询)：<input name="SQL_Query" value="" type="text" size="58"> <input type="submit" value=" 执行 "></td></form>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">数据库文件操作：<b>&nbsp;<a href="admincp.asp?action=database&type=Compact">压缩</a></b>(压缩前最好备份一次)&nbsp;&nbsp;|&nbsp;&nbsp;<b><a href="admincp.asp?action=database&type=Backup">备份</a></b>(强烈推荐每日备份一次)</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" style="padding-left:18px;">
<%IF Request.QueryString("type")="sqlquery" Then
	Dim SQL_Query
	SQL_Query=Request.Form("SQL_Query")
	Conn.ExeCute(SQL_Query)
	SQLQueryNums=SQLQueryNums+1
	Response.Write("<a href=""admincp.asp?action=database"">SQL语句执行成功，请点击返回</a>")
ElseIF Request.QueryString("type")="Compact" Then
	Dim FSO,Engine
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		Response.Write("服务器关闭FSO，无法压缩数据库")
	Else
		If FSO.FileExists(Server.Mappath(AccessPath&"/"&AccessFile)) Then
			Response.Write "压缩数据库开始，网站暂停一切用户的前台操作......<br>"
			Conn.Close
			Set Conn=Nothing
			Application.Lock
			FreeApplicationMemory
			Application(CookieName & "_SiteEnable") = 0
			Application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
			Application.UnLock
			Set Engine = CreateObject("JRO.JetEngine")
			Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(AccessPath&"/"&AccessFile), "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath(AccessPath&"/"&AccessFile&".temp")
			FSO.CopyFile Server.Mappath(AccessPath&"/"&AccessFile&".temp"),Server.Mappath(AccessPath&"/"&AccessFile)
			FSO.DeleteFile(Server.Mappath(AccessPath&"/"&AccessFile&".temp"))
			Set FSO = Nothing
			Set Engine = Nothing
			Response.write "压缩数据库完成..."
			Application.Lock
			Application(CookieName & "_SiteEnable") = 1
			Application(CookieName & "_SiteDisbleWhy") = ""
			Application.UnLock
			Response.Write "<br>网站恢复正常访问..."
			Response.Write("<br><a href=""admincp.asp?action=database"">请点击返回</a>")
		End IF
	End If
	Set FSO=Nothing
ElseIF Request.QueryString("type")="Backup" Then
	Response.Write "备份数据库开始，网站暂停一切用户的前台操作......<br>"
	Conn.Close
	Set Conn=Nothing
	Application.Lock
	Application(CookieName & "_SiteEnable") = 0
	application(CookieName & "_SiteDisbleWhy") = "网站暂停中，请稍候几分钟后再来..."
	Application.UnLock
	CopyFiles Server.Mappath(AccessPath&"/"&AccessFile),Server.Mappath(AccessPath&"/"&AccessFile & "_" & DateToStr(Now(),"YmdHIS") &".BAK")
	Response.write "<br>备份完成..."
	Application.Lock
	Application(CookieName & "_SiteEnable") = 1
	Application(CookieName & "_SiteDisbleWhy") = ""
	Application.UnLock
	Response.write "<br>网站恢复正常访问..."
	Response.Write("<br><a href=""admincp.asp?action=database"">请点击返回</a>")
ElseIF Request.QueryString("type")="Restore" Then

ElseIF Request.QueryString("type")="DeleFile" Then
	IF Request.QueryString("filename")=Empty Then
		Response.Write("<a href=""admincp.asp?action=database"">要删除的文件名不能为空，请点击返回</a>")
	Else
		IF DeleteFiles(Server.MapPath(AccessPath&"/"&Request.QueryString("filename")))=1 Then
			Response.Write("<a href=""admincp.asp?action=database"">文件删除成功，请点击返回</a>")
		Else
			Response.Write("<a href=""admincp.asp?action=database"">文件删除失败，请点击返回</a>")
		End IF
	End IF
Else
	Response.Write("<b>备份文件列表</b><br>")
	Dim DataFolder,DataFileList,DataFile,DataFileName
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		Response.Write("服务器关闭FSO，无法查看备份文件列表")
	Else
		Set DataFolder=FSO.GetFolder(Server.MapPath(AccessPath))
		Set DataFileList=DataFolder.Files
		For Each DataFile IN DataFileList
			IF Ubound(Split(DataFile,"."))>=2 Then
				DataFileName=DataFile.Name
				Response.Write("<font color=""#FF0000"">"&DataFileName&"</font>&nbsp;&nbsp;|&nbsp;&nbsp;<b><a href=""blogdata/"&DataFileName&""">下载此文件</a></b>&nbsp;&nbsp;|&nbsp;&nbsp;<b><a href=""admincp.asp?action=database&type=DeleFile&filename="&DataFileName&""">删除此文件</a></b>&nbsp;&nbsp;|&nbsp;&nbsp;<b><a href=""admincp.asp?action=database&type=Restore&filename="&DataFileName&""">从此文件还原数据</a></b><br>")
			End IF
		Next
	End IF
	Set FSO=Nothing
End If

Function CopyFiles(TempSource,TempEnd)
    Dim FSO
    Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		Response.Write("服务器关闭FSO，无法复制文件")
	Else
		If FSO.FileExists(TempEnd) then
		   Response.Write "目标备份文件 <b>" & TempEnd & "</b> 已存在，请先删除!"
		   Set FSO=Nothing
		   Exit Function
		End IF
		IF FSO.FileExists(TempSource) Then
		Else
		   Response.Write "要复制的源数据库文件 <b>"&TempSource&"</b> 不存在!"
		   Set FSO=Nothing
		   Exit Function
		End If
		FSO.CopyFile TempSource,TempEnd
		Response.Write "已经成功复制文件 <b>"&TempSource&"</b> 到 <b>"&TempEnd&"</b>"
	End If
    Set FSO = Nothing
End Function
%></td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="category" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 分类管理</td>
  </tr><%IF Request.QueryString("type")="EditCate" Then%>
  <tr>
    <td align="center" bgcolor="#FFFFFF" height="48"><%
	Dim Edit_CateID,Edit_CateName,Edit_CateOrder,Edit_CateEvery,Edit_CateNums,Edit_CateMoveTo
	Edit_CateNums=0
	Edit_CateID=Split(Request.Form("cate_ID"),",")
	Edit_CateName=Split(Request.Form("cate_Name"),",")
	Edit_CateOrder=Split(Request.Form("cate_Order"),",")
	Edit_CateMoveTo=Split(Request.Form("Edit_CateMoveTo"),",")
	For Each Edit_CateEvery  IN Edit_CateID
		IF Edit_CateMoveTo(Edit_CateNums)<>0 Then
			Conn.ExeCute("UPDATE blog_Content SET log_CateID="&Edit_CateMoveTo(Edit_CateNums)&" WHERE log_CateID="&Edit_CateID(Edit_CateNums)&"")
			SQLQueryNums=SQLQueryNums+1
		End IF
		Conn.Execute("UPDATE blog_Category SET cate_Name='"&CheckStr(Edit_CateName(Edit_CateNums))&"',cate_Order="&Edit_CateOrder(Edit_CateNums)&" WHERE cate_ID="&Edit_CateEvery&"")
		SQLQueryNums=SQLQueryNums+1
		Edit_CateNums=Edit_CateNums+1
	Next
	IF Request.Form("cate_Dele")<>Empty Then
		Conn.Execute("DELETE  FROM blog_Category WHERE cate_ID IN ("&Request.Form("cate_Dele")&")")
		Conn.Execute("DELETE  FROM blog_Content WHERE log_CateID IN ("&Request.Form("cate_Dele")&")")
		SQLQueryNums=SQLQueryNums+2
	End IF
	Dim new_CateName,new_CateOrder
	new_CateName=CheckStr(Request.Form("new_CateName"))
	new_CateOrder=CheckStr(Request.Form("new_CateOrder"))
	IF new_CateName<>Empty AND new_CateOrder<>Empty Then
		Conn.Execute("INSERT INTO blog_Category(cate_Name,cate_Order) VALUES ('"&new_CateName&"',"&new_CateOrder&")")
		SQLQueryNums=SQLQueryNums+1
	End IF
	Application.Lock()
	Application(CookieName&"_blog_Category")=""
	Application.UnLock()
	Response.Write("<a href=""admincp.asp?action=category"">操作成功，请点击返回</a>")%></td></tr>
  <%Else%><form name="edit_Category" method="post" action="admincp.asp?action=category&type=EditCate">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#EFEFEF">
    <td nowrap>删除？</td>
    <td nowrap>名称</td>
    <td nowrap>排序</td>
	<td width="100%">操作</td>
    </tr><%Dim Adm_CategoryListNums,Adm_CategoryNumI
	Adm_CategoryListNums=Ubound(Arr_Category,2)
	For Adm_CategoryNumI=0 To Adm_CategoryListNums
		Response.Write("<tr bgcolor=""#FFFFFF""><td align=""center""><input name=""cate_Dele"" type=""checkbox"" id=""cate_Dele"" value="""&Arr_Category(0,Adm_CategoryNumI)&"""></td><td><input type=""hidden"" id=""cate_ID"" name=""cate_ID"" value="""&Arr_Category(0,Adm_CategoryNumI)&"""><input type=""text"" size=""15"" id=""cate_Name"" name=""cate_Name"" value="""&Arr_Category(1,Adm_CategoryNumI)&"""></td><td><input type=""text"" size=""5"" id=""cate_Order"" name=""cate_Order"" value="""&Arr_Category(2,Adm_CategoryNumI)&"""></td><td>移动此分类日志到：<select name=""Edit_CateMoveTo"" id=""Edit_CateMoveTo""><option value=""0"">选择分类</option>")
		Dim blog_MoveCateNumI
		For blog_MoveCateNumI=0 To Adm_CategoryListNums
			Response.Write("<option value='"&Arr_Category(0,blog_MoveCateNumI)&"'>"&Arr_Category(1,blog_MoveCateNumI)&"</option>")
		Next
		Response.Write("</select></td></tr>")
	Next
	Response.Write("<tr bgcolor=""#FFFFFF""><td nowrap><b>添加</b>：</td><td><input type=""text"" size=""15"" id=""new_CateName"" name=""new_CateName""></td><td><input type=""text"" size=""5"" id=""new_CateOrder"" name=""new_CateOrder""></td><td></td></tr>")%></table></td>
  </tr><tr>
    <td align="center" bgcolor="#FFFFFF">
      <input type="Submit" name="Submit" value=" 确定编辑 ">
    </td>
  </tr></form><%End IF%>
</table>

<%ElseIF Request.QueryString("action")="forums" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 论坛管理</td>
  </tr><%If Request.QueryString("type")="EditForum" Then%>
  <tr>
    <td align="center" bgcolor="#FFFFFF" height="48"><%
	Dim Edit_ForumID,Edit_ForumName,Edit_ForumOrder,Edit_ForumEvery,Edit_ForumNums,Edit_ForumMoveTo
	Edit_ForumNums=0
	Edit_ForumID=Split(Request.Form("Forum_ID"),",")
	Edit_ForumName=Split(Request.Form("Forum_Name"),",")
	Edit_ForumOrder=Split(Request.Form("Forum_Order"),",")
	Edit_ForumMoveTo=Split(Request.Form("Edit_ForumMoveTo"),",")
	For Each Edit_ForumEvery  IN Edit_ForumID
		If Edit_ForumMoveTo(Edit_ForumNums)<>0 Then
			Conn.ExeCute("UPDATE blog_Threads SET thread_ForumID="&Edit_ForumMoveTo(Edit_ForumNums)&" WHERE thread_ForumID="&Edit_ForumID(Edit_ForumNums)&"")
			Conn.ExeCute("UPDATE blog_Posts SET post_ForumID="&Edit_ForumMoveTo(Edit_ForumNums)&" WHERE post_ForumID="&Edit_ForumID(Edit_ForumNums)&"")
			SQLQueryNums=SQLQueryNums+2
		End If
		Conn.Execute("UPDATE blog_Forums SET forum_Name='"&CheckStr(Edit_ForumName(Edit_ForumNums))&"',forum_Order="&Edit_ForumOrder(Edit_ForumNums)&" WHERE forum_ID="&Edit_ForumEvery&"")
		SQLQueryNums=SQLQueryNums+1
		Edit_ForumNums=Edit_ForumNums+1
	Next
	If Request.Form("Forum_Dele")<>Empty Then
		Conn.Execute("DELETE  FROM blog_Forums WHERE forum_ID IN ("&Request.Form("Forum_Dele")&")")
		Conn.Execute("DELETE  FROM blog_Threads WHERE thread_ForumID IN ("&Request.Form("Forum_Dele")&")")
		Conn.Execute("DELETE  FROM blog_Posts WHERE post_ForumID IN ("&Request.Form("Forum_Dele")&")")
		SQLQueryNums=SQLQueryNums+3
	End If
	Dim new_ForumName,new_ForumOrder
	new_ForumName=CheckStr(Request.Form("new_ForumName"))
	new_ForumOrder=CheckStr(Request.Form("new_ForumOrder"))
	If new_ForumName<>Empty And new_ForumOrder<>Empty Then
		Conn.Execute("INSERT INTO blog_Forums(forum_Name,forum_Order) VALUES ('"&new_ForumName&"',"&new_ForumOrder&")")
		SQLQueryNums=SQLQueryNums+1
	End If
	Application.Lock()
	Application(CookieName&"_blog_Forums")=""
	Application.UnLock()
	Response.Write("<a href=""admincp.asp?action=forums"">操作成功，请点击返回</a>")%></td></tr>
  <%Else%><form name="edit_Forum" method="post" action="admincp.asp?action=forums&type=EditForum">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#EFEFEF">
    <td nowrap>删除？</td>
    <td nowrap>名称</td>
    <td nowrap>排序</td>
	<td width="100%">操作</td>
    </tr><%Dim Adm_ForumsListNums,Adm_ForumsNumI
	Adm_ForumsListNums=Ubound(Arr_Forums,2)
	For Adm_ForumsNumI=0 To Adm_ForumsListNums
		Response.Write("<tr bgcolor=""#FFFFFF""><td align=""center""><input name=""Forum_Dele"" type=""checkbox"" id=""Forum_Dele"" value="""&Arr_Forums(0,Adm_ForumsNumI)&"""></td><td><input type=""hidden"" id=""Forum_ID"" name=""Forum_ID"" value="""&Arr_Forums(0,Adm_ForumsNumI)&"""><input type=""text"" size=""15"" id=""Forum_Name"" name=""Forum_Name"" value="""&Arr_Forums(1,Adm_ForumsNumI)&"""></td><td><input type=""text"" size=""5"" id=""Forum_Order"" name=""Forum_Order"" value="""&Arr_Forums(2,Adm_ForumsNumI)&"""></td><td>移动此论坛帖子到：<select name=""Edit_ForumMoveTo"" id=""Edit_ForumMoveTo""><option value=""0"">选择论坛</option>")
		Dim blog_MoveForumNumI
		For blog_MoveForumNumI=0 To Adm_ForumsListNums
			Response.Write("<option value='"&Arr_Forums(0,blog_MoveForumNumI)&"'>"&Arr_Forums(1,blog_MoveForumNumI)&"</option>")
		Next
		Response.Write("</select></td></tr>")
	Next
	Response.Write("<tr bgcolor=""#FFFFFF""><td nowrap><b>添加</b>：</td><td><input type=""text"" size=""15"" id=""new_ForumName"" name=""new_ForumName""></td><td><input type=""text"" size=""5"" id=""new_ForumOrder"" name=""new_ForumOrder""></td><td></td></tr>")%></table></td>
  </tr><tr>
    <td align="center" bgcolor="#FFFFFF">
      <input type="Submit" name="Submit" value=" 确定编辑 ">
    </td>
  </tr></form><%End If%>
</table>

<%ElseIF Request.QueryString("action")="photocate" Then%>
<table width="100%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
	<td bgcolor="#FFFFFF" class="siderbar_head">相册分类管理</td>
  </tr>
  <%IF Request.QueryString("type")="phCate" Then%>
  <tr>
	<td align="center" bgcolor="#FFFFFF" height="48">
		<%Dim ph_CateID,ph_CateName,ph_CateOrder,ph_CateImage,ph_CateEvery,ph_CateNums,ph_CateMoveTo
		ph_CateNums=0
		ph_CateID=Split(Request.Form("cate_ID"),",")
		ph_CateName=Split(Request.Form("cate_Name"),",")
		ph_CateOrder=Split(Request.Form("cate_Order"),",")
		ph_CateImage=Split(Request.Form("cate_Image"),",")
		ph_CateMoveTo=Split(Request.Form("ph_CateMoveTo"),",")
		For Each ph_CateEvery  IN ph_CateID
			IF ph_CateMoveTo(ph_CateNums)<>0 Then
				Conn.ExeCute("UPDATE photo SET ph_CateID="&ph_CateMoveTo(ph_CateNums)&" WHERE ph_CateID="&ph_CateID(ph_CateNums)&"")
				SQLQueryNums=SQLQueryNums+1
			End IF
			Conn.Execute("UPDATE photo_Cate SET cate_Name='"&CheckStr(ph_CateName(ph_CateNums))&"',cate_Order="&ph_CateOrder(ph_CateNums)&",cate_Image='"&CheckStr(ph_CateImage(ph_CateNums))&"' WHERE cate_ID="&ph_CateEvery&"")
			SQLQueryNums=SQLQueryNums+1
			ph_CateNums=ph_CateNums+1
		Next
		IF Request.Form("cate_Dele")<>Empty Then
			Conn.Execute("DELETE  FROM photo_Cate WHERE cate_ID IN ("&Request.Form("cate_Dele")&")")
			Conn.Execute("DELETE  FROM photo WHERE ph_CateID IN ("&Request.Form("cate_Dele")&")")
			SQLQueryNums=SQLQueryNums+2
		End IF
		Dim ph_new_CateName,ph_new_CateOrder,ph_new_CateImage
		ph_new_CateName=CheckStr(Request.Form("new_CateName"))
		ph_new_CateOrder=CheckStr(Request.Form("new_CateOrder"))
		ph_new_CateImage=CheckStr(Request.Form("new_CateImage"))
		IF ph_new_CateName<>Empty AND ph_new_CateOrder<>Empty Then
			Conn.Execute("INSERT INTO photo_Cate(cate_Name,cate_Order,cate_Image) VALUES ('"&ph_new_CateName&"',"&ph_new_CateOrder&",'"&ph_new_CateImage&"')")
			SQLQueryNums=SQLQueryNums+1
		End IF
		Application.Lock()
		Application(CookieName&"_photo_Cate")=""
		Application.UnLock()%>
		<a href="admincp.asp?action=photocate">操作成功，请点击返回</a>
	</td>
  </tr>
  <%Else%>
	<form action="admincp.asp?action=photocate&type=phCate" method="post" name="ph_Category" id="ph_Category">
  <tr>
	<td align="center" valign="top" bgcolor="#FFFFFF">
		<table width="100%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
		  <tr>
			<td nowrap="nowrap">删除？</td>
			<td nowrap="nowrap">名称</td>
			<td nowrap="nowrap">排序</td>
			<td nowrap="nowrap">图标</td>
			<td width="100%">操作</td>
		  </tr>
		<%Dim Arr_ph_Cate
		If Not IsArray(Application(CookieName&"_photo_Cate")) Then
			Dim ph_CategoryList
			Set ph_CategoryList=Server.CreateObject("ADODB.RecordSet")
			SQL="SELECT cate_ID,cate_Name,cate_Order,cate_Image,cate_Nums FROM photo_Cate ORDER BY cate_Order ASC"
			ph_CategoryList.Open SQL,Conn,1,1
			SQLQueryNums=SQLQueryNums+1
			If ph_CategoryList.EOF And ph_CategoryList.BOF Then
				Redim Arr_ph_Cate(3,0)
			Else
				Arr_ph_Cate=ph_CategoryList.GetRows
			End If
			ph_CategoryList.Close
			Set ph_CategoryList=Nothing
			Application.Lock
			Application(CookieName&"_photo_Cate")=Arr_ph_Cate
			Application.UnLock
		Else
			Arr_ph_Cate=Application(CookieName&"_photo_Cate")
		End IF

		Dim ph_CategoryListNums,ph_CategoryNumI
		ph_CategoryListNums=Ubound(Arr_ph_Cate,2)
		For ph_CategoryNumI=0 To ph_CategoryListNums%>
		  <tr bgcolor="#FFFFFF">
			<td align="center"><input name="cate_Dele" type="checkbox" id="cate_Dele" value=<%=""&Arr_ph_Cate(0,ph_CategoryNumI)&""%>></td>
			<td>
				<input type="hidden" id="cate_ID" name="cate_ID" value=<%=""&Arr_ph_Cate(0,ph_CategoryNumI)&""%>>
				<input type="text" size="15" id="cate_Name" name="cate_Name" value=<%=""&Arr_ph_Cate(1,ph_CategoryNumI)&""%>>
			</td>
			<td><input type="text" size="5" id="cate_Order" name="cate_Order" value=<%=""&Arr_ph_Cate(2,ph_CategoryNumI)&""%>></td>
			<td><input type="text" size="30" id="cate_Image" name="cate_Image" value=<%=""&Arr_ph_Cate(3,ph_CategoryNumI)&""%>></td>
			<td>移动此分类图片到：<select name="ph_CateMoveTo" id="ph_CateMoveTo"><option value="0">选择分类</option>")
			<%Dim ph_MoveCateNumI
			For ph_MoveCateNumI=0 To ph_CategoryListNums
				Response.Write("<option value='"&Arr_ph_Cate(0,ph_MoveCateNumI)&"'>"&Arr_ph_Cate(1,ph_MoveCateNumI)&"</option>")
			Next%>
			</select>
			</td>
		  </tr>
		<%Next%>
		  <tr bgcolor="#FFFFFF">
			<td nowrap><strong>添加</strong>：</td>
			<td><input type="text" size="15" id="new_CateName" name="new_CateName"></td>
			<td><input type="text" size="5" id="new_CateOrder" name="new_CateOrder"></td>
			<td colspan="2" align="left"><input type="text" size="30" id="new_CateImage" name="new_CateImage"></td>
		  </tr>
		</table>
	</td>
  </tr>
  <tr>
	<td align="center" bgcolor="#FFFFFF"><input type="submit" name="Submit" value=" 确定编辑 " /></td>
  </tr>
  </form>
<%End IF%>
	</td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="wordfilter" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 语言过滤</td>
  </tr><%If Request.QueryString("type")="EditWF" Then%>
  <tr>
    <td align="center" bgcolor="#FFFFFF" height="48"><%
	Dim Edit_WFID,Edit_WFString,Edit_WFReplace
	Edit_WFID=Request.Form("wf_ID")
	Edit_WFString=Request.Form("wf_String")
	Edit_WFReplace=Request.Form("wf_Replace")
	Conn.Execute("UPDATE blog_WordFilter SET wf_String='"&CheckStr(Edit_WFString)&"',wf_Replace='"&CheckStr(Edit_WFReplace)&"' WHERE wf_ID="&Edit_WFID&"")
	SQLQueryNums=SQLQueryNums+1
	If Request.Form("wf_Dele")="True" Then
		Conn.Execute("DELETE  FROM blog_WordFilter WHERE wf_ID = "&Edit_WFID)
		SQLQueryNums=SQLQueryNums+1
	End If
	Application.Lock()
	Application(CookieName&"_blog_WordFilter")=""
	Application.UnLock()
	Response.Write("<a href=""admincp.asp?action=wordfilter"">操作成功，请点击返回</a>")
ElseIf Request.QueryString("type")="AddWF" Then
	Dim new_WFString,new_WFReplace
	new_WFString=CheckStr(Request.Form("new_WFString"))
	new_WFReplace=CheckStr(Request.Form("new_WFReplace"))%><tr>
    <td align="center" bgcolor="#FFFFFF" height="48">
  	<%If new_WFString<>Empty And new_WFReplace<>Empty Then
		If Conn.Execute("SELECT COUNT(wf_ID) FROM blog_WordFilter WHERE wf_String='"&new_WFString&"'")(0)=0 Then
			Conn.Execute("INSERT INTO blog_WordFilter(wf_String,wf_Replace) VALUES ('"&new_WFString&"','"&new_WFReplace&"')")
			SQLQueryNums=SQLQueryNums+1
			Application.Lock()
			Application(CookieName&"_blog_WordFilter")=""
			Application.UnLock()
		End If
	End If
	Response.Write("<a href=""admincp.asp?action=wordfilter"">操作成功，请点击返回</a>")%></td></tr>
  <%Else%>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#EFEFEF">
    <td nowrap>删除？</td>
    <td nowrap>字符</td>
	<td nowrap>替换</td>
    <td width="100%">操作</td>
  </tr><%Dim WFRS
  Set WFRS=Server.CreateObject("Adodb.Recordset")
  SQL="SELECT * FROM blog_WordFilter ORDER BY wf_ID DESC"
  WFRS.Open SQL,Conn,1,1
  If Not(WFRS.Eof And WFRS.Bof) Then
  	IF CheckStr(Request.QueryString("Page"))<>Empty Then
				Curpage=Cint(CheckStr(Request.QueryString("Page")))
				IF Curpage<0 Then Curpage=1
			Else
				Curpage=1
			End IF
  	Url_Add="?action=wordfilter&"
	WFRS.PageSize=15
	WFRS.AbsolutePage=CurPage
	MultiPages="<span class=""smalltxt"">"&MultiPage(WFRS.RecordCount,15,CurPage,Url_Add)&"</span>"
  	Response.Write("<tr bgcolor=""#FFFFFF""><td colspan=""4"" nowrap>"&MultiPages&"</td></tr>")
	Do Until WFRS.Eof Or PageCount=15
		Response.Write("<form name=""edit_WordFilter"" method=""post"" action=""admincp.asp?action=wordfilter&type=EditWF""><tr bgcolor=""#FFFFFF""><td align=""center""><input name=""wf_Dele"" type=""checkbox"" id=""wf_Dele"" value=""True""></td><td><input type=""hidden"" id=""wf_ID"" name=""wf_ID"" value="""&WFRS("wf_ID")&"""><input type=""text"" size=""25"" id=""wf_String"" name=""wf_String"" value="""&WFRS("wf_String")&"""></td><td><input type=""text"" size=""25"" id=""wf_Replace"" name=""wf_Replace"" value="""&WFRS("wf_Replace")&"""></td><td><input type=""submit"" name=""Submit"" value="" 编 辑 ""></td></tr></form>")
		WFRS.Movenext
		PageCount=PageCount+1
	Loop
	End If
	WFRS.Close
	Set WFRS=Nothing
	Response.Write("<form name=""add_WordFilter"" method=""post"" action=""admincp.asp?action=wordfilter&type=AddWF""><tr bgcolor=""#FFFFFF""><td nowrap><b>添加</b>：</td><td><input type=""text"" size=""25"" id=""new_WFString"" name=""new_WFString""></td><td><input type=""text"" size=""25"" id=""new_WFReplace"" name=""new_WFReplace""></td><td><input type=""submit"" name=""Submit"" value="" 添加 ""></td></tr></form>")%></table></td>
  </tr><%End If%>
</table>

<%ElseIF Request.QueryString("action")="smilies" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 表情管理</td>
  </tr><%IF Request.QueryString("type")="EditSmilies" Then%>
  <tr>
    <td align="center" bgcolor="#FFFFFF" height="48"><%
	Dim Edit_SmiliesID,Edit_SmiliesText,Edit_SmiliesImage,Edit_SmiliesEvery,Edit_SmiliesNums
	Edit_SmiliesNums=0
	Edit_SmiliesID=Split(Request.Form("smilies_ID"),",")
	Edit_SmiliesText=Split(Request.Form("smilies_Text"),",")
	Edit_SmiliesImage=Split(Request.Form("smilies_Image"),",")
	For Each Edit_SmiliesEvery  IN Edit_SmiliesID
		Conn.Execute("UPDATE blog_Smilies SET sm_Text='"&CheckStr(Edit_SmiliesText(Edit_SmiliesNums))&"',sm_Image='"&CheckStr(Edit_SmiliesImage(Edit_SmiliesNums))&"' WHERE sm_ID="&Edit_SmiliesEvery&"")
		SQLQueryNums=SQLQueryNums+1
		Edit_SmiliesNums=Edit_SmiliesNums+1
	Next
	IF Request.Form("smilies_Dele")<>Empty Then
		Conn.Execute("DELETE  FROM blog_Smilies WHERE sm_ID IN ("&Request.Form("smilies_Dele")&")")
		SQLQueryNums=SQLQueryNums+1
	End IF
	Dim new_SmiliesText,new_SmiliesImage
	new_SmiliesText=CheckStr(Request.Form("new_SmiliesText"))
	new_SmiliesImage=CheckStr(Request.Form("new_SmiliesImage"))
	IF new_SmiliesText<>Empty AND new_SmiliesImage<>Empty Then
		Conn.Execute("INSERT INTO blog_Smilies(sm_Text,sm_Image) VALUES ('"&new_SmiliesText&"','"&new_SmiliesImage&"')")
		SQLQueryNums=SQLQueryNums+1
	End IF
	Application.Lock()
	Application(CookieName&"_blog_Smilies")=""
	Application.UnLock()
	Response.Write("<a href=""admincp.asp?action=smilies"">操作成功，请点击返回</a>")%></td></tr>
  <%Else%><form name="edit_Smilies" method="post" action="admincp.asp?action=smilies&type=EditSmilies">
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF"><table width="90%" border="0" cellpadding="4" cellspacing="1" bgcolor="#CCCCCC">
  <tr bgcolor="#EFEFEF">
    <td nowrap>删除？</td>
    <td nowrap>代码</td>
	<td nowrap>名称</td>
    <td width="100%">图片</td>
    </tr><%Dim Adm_SmiliesNums,Adm_SmiliesNumI
	Adm_SmiliesNums=Ubound(Arr_Smilies,2)
	For Adm_SmiliesNumI=0 To Adm_SmiliesNums
		Response.Write("<tr bgcolor=""#FFFFFF""><td align=""center""><input name=""smilies_Dele"" type=""checkbox"" id=""smilies_Dele"" value="""&Arr_Smilies(0,Adm_SmiliesNumI)&"""></td><td><input type=""hidden"" id=""smilies_ID"" name=""smilies_ID"" value="""&Arr_Smilies(0,Adm_SmiliesNumI)&"""><input type=""text"" size=""15"" id=""smilies_Text"" name=""smilies_Text"" value="""&Arr_Smilies(2,Adm_SmiliesNumI)&"""></td><td><input type=""text"" size=""25"" id=""smilies_Image"" name=""smilies_Image"" value="""&Arr_Smilies(1,Adm_SmiliesNumI)&"""></td><td><img src=""images/smilies/"&Arr_Smilies(1,Adm_SmiliesNumI)&""" border=0""></td></tr>")
	Next
	Response.Write("<tr bgcolor=""#FFFFFF""><td nowrap><b>添加</b>：</td><td><input type=""text"" size=""15"" id=""new_SmiliesText"" name=""new_SmiliesText""></td><td><input type=""text"" size=""25"" id=""new_SmiliesImage"" name=""new_SmiliesImage""></td><td></td></tr>")%></table></td>
  </tr><tr>
    <td align="center" bgcolor="#FFFFFF">
      <input type="Submit" name="Submit" value=" 确定编辑 ">
    </td>
  </tr></form><%End IF%>
</table>
<%ElseIF Request.QueryString("action")="keywords" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 关键字管理</td>
  </tr><%IF Request.QueryString("type")="editKeywords" Then
		Dim Edit_KeywordsID,Edit_KeywordsText,Edit_KeywordsURL,Edit_KeywordsImage,Edit_KeywordsEvery
		Edit_KeywordsID=CheckStr(Request.Form("Keywords_ID"))
		Edit_KeywordsText=CheckStr(Request.Form("Keywords_Text"))
		Edit_KeywordsURL=CheckStr(Request.Form("Keywords_URL"))
		Edit_KeywordsImage=CheckStr(Request.Form("Keywords_Image"))
		Conn.Execute("UPDATE blog_Keywords SET key_Text='"&Edit_KeywordsText&"',key_URL='"&Edit_KeywordsURL&"',key_Image='"&Edit_KeywordsImage&"' WHERE key_ID="&Edit_KeywordsID&"")
		SQLQueryNums=SQLQueryNums+1
		IF Request.Form("Keywords_Dele")<>Empty Then
			Conn.Execute("DELETE  FROM blog_Keywords WHERE key_ID="&Request.Form("Keywords_Dele"))
			SQLQueryNums=SQLQueryNums+1
		End IF
		Application.Lock()
		Application(CookieName&"_blog_Keywords")=""
		Application.UnLock()
		Response.Write("<tr><td align=""center"" bgcolor=""#FFFFFF"" height=""48""><a href=""admincp.asp?action=keywords"">操作成功，请点击返回</a></td></tr>")
	ElseIF Request.QueryString("type")="addKeywords" Then
		Dim new_KeywordsText,new_KeywordsURL,new_KeywordsImage
		new_KeywordsText=CheckStr(Request.Form("new_KeywordsText"))
		new_KeywordsURL=CheckStr(Request.Form("new_KeywordsURL"))
		new_KeywordsImage=CheckStr(Request.Form("new_KeywordsImage"))
		IF new_KeywordsText<>Empty AND new_KeywordsURL<>Empty Then
			Conn.Execute("INSERT INTO blog_Keywords(key_Text,key_URL,key_Image) VALUES ('"&new_KeywordsText&"','"&new_KeywordsURL&"','"&new_KeywordsImage&"')")
			SQLQueryNums=SQLQueryNums+1
		End IF
		Application.Lock()
		Application(CookieName&"_blog_Keywords")=""
		Application.UnLock()
		Response.Write("<tr><td align=""center"" bgcolor=""#FFFFFF"" height=""48""><a href=""admincp.asp?action=keywords"">操作成功，请点击返回</a></td></tr>")
		Else%>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF">
	<div style="overflow-y: scroll;overflow-x:hidden;height: 400px;"><table width="98%" border="0" cellpadding="4" cellspacing="1" bgcolor="#DFDFDF">
  <tr bgcolor="#EFEFEF">
	<td nowrap>删除？</td>
    <td nowrap>文本</td>
	<td nowrap>链接</td>
	<td nowrap>图片</td>
    <td nowrap>图片预览</td>
	<td width="100%">操作</td>
    </tr><%Dim Adm_KeywordsNums,Adm_KeywordsNumI
	Adm_KeywordsNums=Ubound(Arr_Keywords,2)
	For Adm_KeywordsNumI=0 To Adm_KeywordsNums
		Response.Write("<form name=""edit_Keywords"" method=""post"" action=""admincp.asp?action=keywords&type=editKeywords""><tr bgcolor=""#FFFFFF""><td align=""center""><input name=""Keywords_Dele"" type=""checkbox"" id=""Keywords_Dele"" value="""&Arr_Keywords(0,Adm_KeywordsNumI)&"""></td><td><input type=""hidden"" id=""Keywords_ID"" name=""Keywords_ID"" value="""&Arr_Keywords(0,Adm_KeywordsNumI)&"""><input type=""text"" size=""10"" id=""Keywords_Text"" name=""Keywords_Text"" value="""&Arr_Keywords(1,Adm_KeywordsNumI)&"""></td><td><input type=""text"" size=""38"" id=""Keywords_URL"" name=""Keywords_URL"" value="""&Arr_Keywords(2,Adm_KeywordsNumI)&"""></td>")
		IF Arr_Keywords(3,Adm_KeywordsNumI)=Empty Then
			Response.Write("<td><input type=""text"" size=""15"" id=""Keywords_Image"" name=""Keywords_Image"" value=""""></td><td></td>")
		Else
			Response.Write("<td><input type=""text"" size=""15"" id=""Keywords_Image"" name=""Keywords_Image"" value="""&Arr_Keywords(3,Adm_KeywordsNumI)&"""></td><td><img src=""images/keywords/"&Arr_Keywords(3,Adm_KeywordsNumI)&""" border=0""></td>")
		End IF
		Response.Write("<td><input type=""Submit"" name=""Submit"" value="" 编辑 ""></td></tr></form>")
	Next
	%></table></div>
	<%Response.Write("<table width=""98%"" border=""0"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#DFDFDF""><form name=""add_Keywords"" method=""post"" action=""admincp.asp?action=keywords&type=addKeywords""><tr bgcolor=""#FFFFFF""><td>文本：<input type=""text"" size=""10"" id=""new_KeywordsText"" name=""new_KeywordsText""></td><td>链接：<input type=""text"" size=""38"" id=""new_KeywordsURL"" name=""new_KeywordsURL""></td><td>图片：<input type=""text"" size=""15"" id=""new_KeywordsImage"" name=""new_KeywordsImage""></td><td><input type=""Submit"" name=""Submit"" value="" 添加 ""></td></tr></form></table>")%></td>
  </tr><%End IF%>
</table>
<%ElseIF Request.QueryString("action")="favorite" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 网络书签管理</td>
  </tr><%IF Request.QueryString("type")="delete" Then
		Dim DEL_favID
		DEL_favID=CheckStr(Request.QueryString("favID"))
		If DEL_favID=Empty Then
			Response.Write("<tr><td align=""center"" bgcolor=""#FFFFFF"" height=""48""><a href=""admincp.asp?action=favID"">参数错误，请点击返回</a></td></tr>")
		Else
			Conn.Execute("DELETE FROM blog_Favorite WHERE fav_ID="&DEL_favID)
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<tr><td align=""center"" bgcolor=""#FFFFFF"" height=""48""><a href=""admincp.asp?action=favID"">操作成功，请点击返回</a></td></tr>")
		End If
	Else%>
  <tr>
    <td align="center" valign="top" bgcolor="#FFFFFF">
	<div style="overflow-y: scroll;overflow-x:hidden;height: 400px;"><table width="98%" border="0" cellpadding="4" cellspacing="1" bgcolor="#DFDFDF">
  <tr bgcolor="#EFEFEF">
	<td nowrap>ID</td>
    <td nowrap>名称</td>
	<td nowrap>作者</td>
	<td nowrap>链接</td>
    <td nowrap>介绍</td>
	<td width="100%">操作</td>
    </tr><%Dim fav_AdmList
	Set fav_AdmList = Conn.ExeCute("SELECT * FROM blog_Favorite ORDER BY fav_PostTime ASC")
	Do While Not fav_AdmList.Eof
		Response.Write("<tr bgcolor=""#FFFFFF""><td align=""center"">"&fav_AdmList("fav_ID")&"</td><td nowrap>"&HTMLEncode(fav_AdmList("fav_Name"))&"</td><td nowrap>"&fav_AdmList("fav_Author")&"</td><td><a href="""&CheckLinkStr(fav_AdmList("fav_URL"))&""" target=""_blank""><img src=""images/icon_trackback.gif"" align=""absmiddle"" border=""0"" alt=""点击访问""></a></td><td><img src=""images/sider_other.gif"" align=""absmiddle"" border=""0"" alt="""&HTMLEncode(fav_AdmList("fav_Info"))&"""></td><td><a href=""admincp.asp?action=favorite&type=delete&favID="&fav_AdmList("fav_ID")&""">删除</a></td></tr>")
		fav_AdmList.MoveNext
	Loop
	Set fav_AdmList = Nothing
	%></table></div></td>
  </tr><%End IF%>
</table>
<%ElseIF Request.QueryString("action")="setting" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 一般设置</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" align="center" height="48"><%IF Request.QueryString("type")="EnableSite" Then
		Application.Lock()
		Application(CookieName & "_SiteEnable") = 1
		Application(CookieName & "_SiteDisbleWhy") = ""
		Application.UnLock()
		Response.Write("<a href=""admincp.asp?action=setting"">开启站点成功，请点击返回</a>")
	ElseIF Request.QueryString("type")="DisableSite" Then
		Set Conn=Nothing
		FreeApplicationMemory
		Application.Lock()
		Application(CookieName & "_SiteEnable") = 0
		Application(CookieName & "_SiteDisbleWhy")="站点维护中，请稍候再来..."
		Application.UnLock()
		Response.Write("<br><a href=""admincp.asp?action=setting"">关闭站点成功</a>")
	Else
		IF Application(CookieName & "_SiteEnable") = 0 AND Application(CookieName & "_SiteDisbleWhy")<>"" Then
			Response.Write("<b>站点已关闭</b>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admincp.asp?action=setting&type=EnableSite"">点击开启站点</a>")
		Else
			Response.Write("<b>站点已开启</b>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admincp.asp?action=setting&type=DisableSite"">点击关闭站点</a>")
		End IF
	End IF%></td>
  </tr>
</table><br><table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 基本信息</td>
  </tr>
  <tr>
    <%IF Request.QueryString("type")="baseInfo" Then
		Dim edit_blogName,edit_blogURL,edit_blogPerpage,edit_blogSpamIP,edit_threadPerpage,edit_postPerpage
		edit_blogName=CheckStr(Request.Form("blog_Name"))
		edit_blogURL=CheckStr(Request.Form("blog_URL"))
		edit_blogPerpage=CheckStr(Request.Form("blog_PerPage"))
		edit_blogSpamIP=CheckStr(Request.Form("spam_IP"))
		edit_threadPerpage=CheckStr(Request.Form("thread_PerPage"))
		edit_postPerpage=CheckStr(Request.Form("post_PerPage"))
		IF edit_blogName<>Empty OR edit_blogURL<>Empty OR IsInteger(edit_blogPerpage)=True Then
			Conn.ExeCute("UPDATE blog_Info SET blog_Name='"&edit_blogName&"',blog_URL='"&edit_blogURL&"',thread_PerPage='"&edit_threadPerpage&"',post_PerPage='"&edit_postPerpage&"',blog_PerPage="&edit_blogPerpage&",blog_SpamIP='"&edit_blogSpamIP&"'")
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<td bgcolor=""#FFFFFF"" align=""center"" height=""48""><a href=""admincp.asp?action=setting"">基本信息修改成功，请点击返回</a></td>")
		Else
			Response.Write("<td bgcolor=""#FFFFFF"" align=""center"" height=""48""><a href=""admincp.asp?action=setting"">基本信息修改失败，请点击返回</a></td>")
		End IF
	Else
		Response.Write("<form name=""blogInfo"" action=""admincp.asp?action=setting&type=baseInfo"" method=""post""><td bgcolor=""#FFFFFF"" align=""center"" height=""48"">BLOG 名称：<input name=""blog_Name"" value="""&siteName&""" type=""text"">&nbsp;&nbsp;&nbsp;BLOG 地址：<input name=""blog_URL"" value="""&siteURL&""" type=""text"" size=""25""><br>&nbsp;&nbsp;&nbsp;每页日志：<input name=""blog_PerPage"" size=""2"" type=""text"" value="""&blogPerPage&""">&nbsp;&nbsp;&nbsp;每页主题：<input name=""thread_PerPage"" size=""2"" type=""text"" value="""&threadPerPage&""">&nbsp;&nbsp;&nbsp;每页帖子：<input name=""post_PerPage"" size=""2"" type=""text"" value="""&postPerPage&"""><br>IP禁止：<textarea id=""spam_IP"" name=""spam_IP"" cols=""48"" rows=""5"">"&blog_SpamIP&"</textarea><br /><br /><center><input type=""submit"" value="" 确定 ""></td></form>")
	End IF%>
  </tr>
</table><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 数据统计</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" align="center" height="48"><%IF Request.QueryString("type")="memPosts" Then
		Dim PostNums_MemList
		SET PostNums_MemList=Server.CreateObject("Adodb.Recordset")
		SQL="SELECT mem_Name,mem_PostLogs,mem_PostComms FROM blog_Member"
		PostNums_MemList.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		Do While Not PostNums_MemList.EOF
			PostNums_MemList("mem_PostLogs")=Conn.ExeCute("SELECT COUNT(log_ID) FROM blog_Content WHERE log_Author="""&PostNums_MemList("mem_Name")&"""")(0)
			SQLQueryNums=SQLQueryNums+1
			PostNums_MemList("mem_PostComms")=Conn.ExeCute("SELECT COUNT(comm_ID) FROM blog_Comment WHERE comm_Author="""&PostNums_MemList("mem_Name")&"""")(0)
			SQLQueryNums=SQLQueryNums+1
			PostNums_MemList.Update
			PostNums_MemList.MoveNext
		Loop
		PostNums_MemList.Close
		SET PostNums_MemList=Nothing
		Response.Write("<a href=""admincp.asp?action=setting"">统计用户发表数成功，请点击返回</a>")
	ElseIF Request.QueryString("type")="blogNums" Then
		Dim blog_Info
		SET blog_Info=Server.CreateObject("Adodb.Recordset")
		SQL="SELECT blog_LogNums,blog_CommNums,blog_MemNums,blog_downloadNums,blog_GuestbookNums,blog_VisitNums,blog_QuoteNums,blog_ThreadNums,blog_PostNums,blog_photoNums FROM blog_Info"
		blog_Info.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		Do While Not blog_Info.EOF
			blog_Info("blog_CommNums")=Conn.ExeCute("SELECT COUNT(comm_ID) FROM blog_Comment")(0)
			blog_Info("blog_LogNums")=Conn.ExeCute("SELECT COUNT(log_ID) FROM blog_Content")(0)
			blog_Info("blog_MemNums")=Conn.ExeCute("SELECT COUNT(mem_ID) FROM blog_Member")(0)
			blog_Info("blog_GuestbookNums")=Conn.ExeCute("SELECT COUNT(gb_ID) FROM blog_Guestbook")(0)
			blog_Info("blog_downloadNums")=Conn.ExeCute("SELECT COUNT(downl_ID) FROM blog_Download")(0)
			blog_Info("blog_VisitNums")=Conn.ExeCute("SELECT COUNT(coun_ID) FROM blog_Counter")(0)
			blog_Info("blog_QuoteNums")=Conn.ExeCute("SELECT COUNT(tb_ID) FROM blog_TrackBack")(0)
			blog_Info("blog_ThreadNums")=Conn.ExeCute("SELECT COUNT(thread_ID) FROM blog_Threads")(0)
			blog_Info("blog_PostNums")=Conn.ExeCute("SELECT COUNT(post_ID) FROM blog_Posts")(0)
			blog_Info("blog_photoNums")=Conn.ExeCute("SELECT COUNT(photo_ID) FROM blog_photo")(0)
			SQLQueryNums=SQLQueryNums+10
			blog_Info.Update
			blog_Info.MoveNext
		Loop
		blog_Info.Close
		SET blog_Info=Nothing
		Response.Write("<a href=""admincp.asp?action=setting"">统计BLOG数据成功，请点击返回</a>")
	ElseIF Request.QueryString("type")="blogPosts" Then
		Dim Nums_CommList
		Set Nums_CommList=Server.CreateObject("Adodb.Recordset")
		SQL="SELECT log_ID,log_CommNums,log_QuoteNums FROM blog_Content"
		Nums_CommList.Open SQL,Conn,1,3
		SQLQueryNums=SQLQueryNums+1
		Do While Not Nums_CommList.EOF
			Nums_CommList("log_CommNums")=Conn.ExeCute("SELECT COUNT(comm_ID) FROM blog_Comment WHERE blog_ID="&Nums_CommList("log_ID"))(0)
			Nums_CommList("log_QuoteNums")=Conn.ExeCute("SELECT COUNT(tb_ID) FROM blog_TrackBack WHERE blog_ID="&Nums_CommList("log_ID"))(0)
			SQLQueryNums=SQLQueryNums+2
			Nums_CommList.Update
			Nums_CommList.MoveNext
		Loop
		Nums_CommList.Close
		SET Nums_CommList=Nothing
		Response.Write("<a href=""admincp.asp?action=setting"">统计日志评论数成功，请点击返回</a>")
	ElseIf Request.QueryString("type")="blogVisitNums" Then
		Conn.ExeCute("UPDATE blog_Info SET blog_VisitBaseNums="&blog_VisitNums&",blog_VisitNums=0")
		Conn.ExeCute("DELETE  FROM blog_Counter")
		SQLQueryNums=SQLQueryNums+2
		Response.Write("<a href=""admincp.asp?action=setting"">清空访问记录表成功，请点击返回</a>")
	Else
		Response.Write("<b><a href=""admincp.asp?action=setting&type=blogNums"">统计BLOG数据</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admincp.asp?action=setting&type=memPosts"">统计用户发表数</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admincp.asp?action=setting&type=blogPosts"">统计日志评论数</a>&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;<a href=""admincp.asp?action=setting&type=blogVisitNums"">清空访问记录表</a>")
	End IF%></td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="attachment" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 附件管理</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><%IF Request.QueryString("type")="DeleFile" Then
		IF Request.QueryString("filename")=Empty Then
			Response.Write("<a href=""admincp.asp?action=attachment"">没有选择要删除的文件，请点击返回</a>")
		Else
			IF DeleteFiles(Server.MapPath("attachments/"&Request.QueryString("filename")))=1 Then
				Response.Write("<a href=""admincp.asp?action=attachment"">文件删除成功，请点击返回</a>")
			Else
				Response.Write("<a href=""admincp.asp?action=attachment"">文件删除失败，请点击返回</a>")
			End IF
		End IF
	Else
		Dim AttachmentFolder
		Set FSO=Server.CreateObject("Scripting.FileSystemObject")
		IF Err<>0 Then
			Err.Clear
			Response.Write("服务器关闭FSO，无法查看附件")
		Else
			If Request.QueryString("FolderName")=Empty Then
				Set AttachmentFolder=FSO.GetFolder(Server.MapPath("attachments"))
				Dim AttachmentFolderList
				Set AttachmentFolderList=AttachmentFolder.SubFolders
				Dim AttachmentFolderName,AttachmentFolderEvery
				For Each AttachmentFolderEvery IN AttachmentFolderList
					AttachmentFolderName=AttachmentFolderEvery.Name
					Response.Write("<a href=""admincp.asp?action=attachment&foldername="&AttachmentFolderName&""">浏览文件夹"&AttachmentFolderName&"中的附件</a><br><img name=""HideImage"" src="""" width=""2"" height=""5"" alt="""" style=""background-color: #FFFFFF""><br>")
				Next
				Set AttachmentFolderList=Nothing
			Else
				Set AttachmentFolder=FSO.GetFolder(Server.MapPath("attachments/"&Request.QueryString("FolderName")))
				Dim AttachmentFileList
				Set AttachmentFileList=AttachmentFolder.Files
				Dim AttachmentFileName,AttachmentFileEvery
				For Each AttachmentFileEvery IN AttachmentFileList
					AttachmentFileName=AttachmentFileEvery.Name
					Response.Write("浏览附件 <a href=""attachments/"&Request.QueryString("FolderName")&"/"&AttachmentFileName&""" target=""_blank"">"&AttachmentFileName&"</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""admincp.asp?action=attachment&type=DeleFile&filename="&Request.QueryString("FolderName")&"/"&AttachmentFileName&"""><b>删除文件</b></a><br><img name=""HideImage"" src="""" width=""2"" height=""5"" alt="""" style=""background-color: #FFFFFF""><br>")
				Next
				Set AttachmentFileList=Nothing
			End If
			Set AttachmentFolder=Nothing
		End If
		Set FSO=Nothing
	End IF%></td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="linkscheck" Then%>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 友情链接验证</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><%IF Request.QueryString("type")="linkcheckedit" Then
		IF CheckStr(Request.QueryString("linkID"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">参数错误，点击返回上一页</a>")
		Else
			Conn.ExeCute("UPDATE blog_Links SET link_IsShow=True WHERE link_ID="&CheckStr(Request.QueryString("linkID"))&"")
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<a href=""admincp.asp?action=linkscheck"">链接验证成功，点击返回</a>")
		End IF
	ElseIF Request.QueryString("type")="linkcheckdele" Then
		IF CheckStr(Request.QueryString("linkID"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">参数错误，点击返回上一页</a>")
		Else
			Conn.ExeCute("DELETE  FROM blog_Links WHERE link_ID="&CheckStr(Request.QueryString("linkID"))&"")
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<a href=""admincp.asp?action=linkscheck"">链接删除成功，点击返回</a>")
		End IF
	Else
		Dim blog_LinksCheck
		Set blog_LinksCheck=Conn.Execute("SELECT * FROM blog_Links WHERE link_IsShow=False ORDER BY link_Order ASC")
		SQLQueryNums=SQLQueryNums+1
		IF blog_LinksCheck.EOF AND blog_LinksCheck.BOF Then
			Response.Write("暂时没有需要验证的友情链接")
		Else
			Response.Write("<div style=""overflow-y: scroll;overflow-x:hidden;height: 388px;""><table border=""0"" cellpadding=""3"" cellspacing=""1"" bgcolor=""#DFDFDF"" width=""98%""><tr bgcolor=""#EFEFEF""><td><b>需要验证的友情链接</b></td><td><b>操作</b></td></tr>")
			Do While Not blog_LinksCheck.EOF
				Dim blog_LinksCheckImage
				IF blog_LinksCheck("link_Image")<>Empty Then
					blog_LinksCheckImage="<img src='"&CheckLinkStr(CheckLinkStr(blog_LinksCheck("link_Image")))&"' border='0' />"
				Else
					blog_LinksCheckImage="没有图片"
				End IF
				Response.Write("<tr bgcolor=""#FFFFFF""><td><a href="""&CheckLinkStr(blog_LinksCheck("link_URL"))&""" alt="""&CheckLinkStr(blog_LinksCheckImage)&""" target=""_blank"">"&HTMLEncode(blog_LinksCheck("link_Name"))&"</a></td><td><a href=""admincp.asp?action=linkscheck&type=linkcheckedit&linkID="&blog_LinksCheck("link_ID")&""">通过链接验证</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""admincp.asp?action=linkscheck&type=linkcheckdele&linkID="&blog_LinksCheck("link_ID")&""">删除此链接</a></td></tr>")
				blog_LinksCheck.MoveNext
			Loop
			Response.Write("</table></div>")
		End IF
		Set blog_LinksCheck=Nothing
	End IF%></td>
  </tr>
</table>
<%ElseIF Request.QueryString("action")="links" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 友情链接管理</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><%IF Request.QueryString("type")="linkedit" Then
		IF CheckStr(Request.QueryString("linkID"))=Empty OR CheckStr(Request.Form("edit_blogLinkName"))=Empty OR CheckStr(Request.Form("edit_blogLinkURL"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">请将信息填写完整，点击返回上一页</a>")
		Else
			IF CheckStr(Request.Form("edit_blogLinkDele"))="True" Then
				Conn.ExeCute("DELETE  FROM blog_Links WHERE link_ID="&CheckStr(Request.QueryString("linkID"))&"")
				SQLQueryNums=SQLQueryNums+1
			Else
				Dim link_IsMainValue
				IF CheckStr(Request.Form("edit_blogLinkMain"))="True" Then
					link_IsMainValue=True
				Else
					link_IsMainValue=False
				End IF
				Conn.ExeCute("UPDATE blog_Links SET link_Name='"&CheckStr(Request.Form("edit_blogLinkName"))&"',link_Image='"&CheckStr(Request.Form("edit_blogLinkImage"))&"',link_URL='"&CheckStr(Request.Form("edit_blogLinkURL"))&"',link_Order="&CheckStr(Request.Form("edit_blogLinkOrder"))&",link_IsMain="&link_IsMainValue&" WHERE link_ID="&CheckStr(Request.QueryString("linkID"))&"")
				SQLQueryNums=SQLQueryNums+1
			End IF
			Application.Lock
			Application(CookieName&"_blog_Bloglinks")=""
			Application.UnLock
			Response.Write("<a href=""admincp.asp?action=links"">编辑友情链接成功，请点击返回</a>")
		End IF
	ElseIF Request.QueryString("type")="linkadd" Then
		IF CheckStr(Request.Form("new_blogLinkName"))=Empty OR CheckStr(Request.Form("new_blogLinkURL"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">请将信息填写完整，点击返回上一页</a>")
		Else
			Conn.ExeCute("INSERT INTO blog_Links(link_Name,link_Image,link_URL,link_Order,link_IsShow) VALUES ('"&CheckStr(Request.Form("new_blogLinkName"))&"','"&CheckStr(Request.Form("new_blogLinkImage"))&"','"&CheckStr(Request.Form("new_blogLinkURL"))&"',"&CheckStr(Request.Form("new_blogLinkOrder"))&",True)")
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<a href=""admincp.asp?action=links"">添加友情链接成功，请点击返回</a>")
		End IF
	Else
		Dim blog_Links
		Set blog_Links=Conn.Execute("SELECT * FROM blog_Links WHERE link_IsShow=True ORDER BY link_Order ASC")
		SQLQueryNums=SQLQueryNums+1
		IF blog_Links.EOF AND blog_Links.BOF Then
			Response.Write("暂时没有友情链接")
		Else
			Response.Write("<div style=""overflow-y: scroll;overflow-x:hidden;height: 388px;""><table border=""0"" cellpadding=""3"" cellspacing=""1"" bgcolor=""#DFDFDF"" width=""98%""><tr bgcolor=""#FFFFFF""><td>删？</td><td>首？</td><td>名称</td><td>图片</td><td>链接</td><td>排序</td><td>操作</td></tr>")
			Do While Not blog_Links.EOF
				Dim link_IsMainChecked
				IF blog_Links("link_IsMain")=True Then
					link_IsMainChecked=" checked"
				Else
					link_IsMainChecked=""
				End IF
				Response.Write("<tr bgcolor=""#FFFFFF""><form name=""addlink"" action=""admincp.asp?action=links&type=linkedit&linkID="&blog_Links("link_ID")&""" method=""post""><td><input name=""edit_blogLinkDele"" type=""checkbox"" id=""edit_blogLinkDele"" value=""True""></td><td><input name=""edit_blogLinkMain"" type=""checkbox"" id=""edit_blogLinkMain"" value=""True"""&link_IsMainChecked&"></td><td><input type=""text"" size=""10"" id=""edit_blogLinkName"" name=""edit_blogLinkName"" value="""&CheckLinkStr(blog_Links("link_Name"))&"""></td><td><input type=""text"" size=""25"" id=""edit_blogLinkImage"" name=""edit_blogLinkImage"" value="""&CheckLinkStr(blog_Links("link_Image"))&"""></td><td><input type=""text"" size=""15"" id=""edit_blogLinkURL"" name=""edit_blogLinkURL"" value="""&CheckLinkStr(blog_Links("link_URL"))&"""></td><td><input type=""text"" size=""2"" id=""edit_blogLinkOrder"" name=""edit_blogLinkOrder"" value="""&CheckLinkStr(blog_Links("link_Order"))&"""></td><td><input type=""Submit"" name=""Submit"" value="" 编 辑 ""></td></form></tr>")
				blog_Links.MoveNext
			Loop
			Response.Write("</table></div>")
		End IF
		Set blog_Links=Nothing
		Response.Write("<hr width=""100%"" size=""1""><table><tr><form name=""addlink"" action=""admincp.asp?action=links&type=linkadd"" method=""post""><td>名称：<input type=""text"" size=""10"" id=""new_blogLinkName"" name=""new_blogLinkName"">&nbsp;&nbsp;图片：<input type='text' size='20' id='new_blogLinkImage' name='new_blogLinkImage'>&nbsp;&nbsp;链接：<input type=""text"" size=""16"" id=""new_blogLinkURL"" name=""new_blogLinkURL"">&nbsp;&nbsp;排序：<input type=""text"" size=""1"" id=""new_blogLinkOrder"" name=""new_blogLinkOrder"">&nbsp;&nbsp;<input type=""Submit"" name=""Submit"" value="" 添  加 ""></td></form></tr></table>")
	End IF%></td>
  </tr>
</table>
<%
ElseIF Request.QueryString("action")="download" Then%>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" class="formbox">
        <%IF Request.QueryString("type")="adddownload" Then%>
        <tr> 
          <td class="siderbar_head" colspan="5">下载</td>
        </tr>
        <%
	Dim Show_downlID,Show_downlName,Show_downlAuthor,Show_downlPostTime
	Show_downlID=Split(Request.Form("sdownl_IDs"),",")
	Show_downlName=Split(Request.Form("sdownl_Names"),",")
	Show_downlAuthor=Split(Request.Form("sdownl_Authors"),",")
	Show_downlPostTime=Split(Request.Form("sdownl_PostTimes"),",")
	IF Request.Form("downl_Dele")<>Empty Then
		Conn.Execute("DELETE  FROM blog_Download WHERE downl_ID IN ("&CheckStr(Request.Form("downl_Dele"))&")")
		Conn.ExeCute("UPDATE blog_Info SET blog_downloadNums=blog_downloadNums-1")
		Dim downlNums
	    set downlNums=Conn.ExeCute("select count(*) as down_Rows FROM blog_Download")
		Conn.Execute("UPDATE blog_Toolbox SET tb_Nums="&downlNums("down_Rows")&" WHERE tb_Name='Download'")
	End IF
	Dim new_downlCate,new_downlName,new_downlAuthor,new_downlFrom,new_downlsize,new_downlFromURL,new_downlImgPath,new_downlComment,new_downlDcomm1,new_downlDcomm1URL,new_downlDcomm2,new_downlDcomm2URL,new_downlDcomm3,new_downlDcomm3URL,new_downlDcomm4,new_downlDcomm4URL
	new_downlCate=CheckStr(Request.Form("new_downlCate"))
	new_downlName=CheckStr(Request.Form("new_downlName"))
	new_downlAuthor=CheckStr(Request.Form("new_downlAuthor"))
	new_downlFrom=CheckStr(Request.Form("new_downlFrom"))
	new_downlsize=CheckStr(Request.Form("new_downlsize"))
	new_downlFromURL=CheckStr(Request.Form("new_downlFromURL"))
	new_downlImgPath=CheckStr(Request.Form("new_downlImgPath"))
	new_downlComment=CheckStr(Request.Form("new_downlComment"))
	new_downlDcomm1=CheckStr(Request.Form("new_downlDcomm1"))
	new_downlDcomm1URL=CheckStr(Request.Form("new_downlDcomm1URL"))
	new_downlDcomm2=CheckStr(Request.Form("new_downlDcomm2"))
	new_downlDcomm2URL=CheckStr(Request.Form("new_downlDcomm2URL"))
	new_downlDcomm3=CheckStr(Request.Form("new_downlDcomm3"))
	new_downlDcomm3URL=CheckStr(Request.Form("new_downlDcomm3URL"))
	new_downlDcomm4=CheckStr(Request.Form("new_downlDcomm4"))
	new_downlDcomm4URL=CheckStr(Request.Form("new_downlDcomm4URL"))
    IF new_downlName<>Empty AND new_downlFrom<>Empty AND new_downlDcomm1URL<>Empty AND new_downlDcomm1<>Empty Then
	Conn.Execute("INSERT INTO blog_Download(downl_Cate,downl_Name,downl_Author,downl_From,downl_size,downl_FromURL,downl_ImgPath,downl_Comment,downl_Dcomm1,downl_Dcomm1URL,downl_Dcomm2,downl_Dcomm2URL,downl_Dcomm3,downl_Dcomm3URL,downl_Dcomm4,downl_Dcomm4URL) VALUES ("&new_downlCate&",'"&new_downlName&"','"&new_downlAuthor&"','"&new_downlFrom&"','"&new_downlsize&"','"&new_downlFromURL&"','"&new_downlImgPath&"','"&new_downlComment&"','"&new_downlDcomm1&"','"&new_downlDcomm1URL&"','"&new_downlDcomm2&"','"&new_downlDcomm2URL&"','"&new_downlDcomm3&"','"&new_downlDcomm3URL&"','"&new_downlDcomm4URL&"','"&new_downlDcomm4URL&"')")
	Conn.ExeCute("UPDATE blog_Toolbox SET tb_Nums=tb_Nums+1 WHERE tb_Name='Download'")
	Conn.ExeCute("UPDATE blog_Info SET blog_downloadNums=blog_downloadNums+1")
	End IF
	Application.Lock()
	Application(CookieName&"_blog_Download")=""
	Application.UnLock()
	Response.Write("<tr><td class='formbox-content' colspan='4' align='center'><br>已保存更改。<br><br><a href='?action=download'>返回</a></td></tr>")
	%>
        <%Else%>
        <form name="input" method="post" action="?action=download&type=adddownload">
          <tr bgcolor="#EFEFEF"> 
            <td width="90" height="24">删除?</td>
            <td width="120">程序名称</td>
            <td width="105" align="center">发布作者</td>
            <td width="120" align="center">发表时间</td>
            <td width="40">操作</td>
          </tr>
          <%Dim Adm_Download,Adm_DownloadContent
	For Each Adm_Download IN Arr_Download
		Adm_DownloadContent=Split(Adm_Download,"$|$")
		%>
          <tr class='formbox-content'> 
            <td align='center'> <input name='downl_Dele' type='checkbox' id='downl_Dele' value='<%=Adm_DownloadContent(0)%>'> 
              <input name='sdownl_IDs' type='hidden' id='sdownl_IDs' value='<%=Adm_DownloadContent(0)%>'> 
            </td>
            <td align='left'> <input type='hidden' size='25' id='sdownl_Names' name='sdownl_Names' value='<%=Adm_DownloadContent(2)%>'> 
              <%=Adm_DownloadContent(2)%></td>
            <td align='left'> <input name='sdownl_Authors' type='hidden' id='sdownl_Authors' value='<%=Adm_DownloadContent(3)%>'> 
              <%=Adm_DownloadContent(3)%> </td>
            <td align="center"> <input name='sdownl_PostTimes' type='hidden' id='sdownl_PostTimes' value='<%=Adm_DownloadContent(7)%>'> 
              <%=Adm_DownloadContent(7)%> </td>
            <td align="center"><a href="downledit.asp?downlID=<%=Adm_DownloadContent(0)%>" target="_blank">编辑</a></td>
          </tr>
          <%Next%>
          <tr class='formbox-title'> 
            <td colspan="5">添加下载</td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">分类:</td>
            <td colspan="3"><select name="new_downlCate">
                <option value="1" selected>文档</option>
                <option value="2">源码</option>
                <option value="3">程序</option>
              </select>
             <font color="#FF0000">*</font>&nbsp;&nbsp;文件名:
              <input type='text' size='30' id='new_downlName' name='new_downlName' maxlength="50">&nbsp;&nbsp;大小:<input type='text' size='5' id='new_downlsize' name='new_downlsize'></td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">来自:</td>
            <td><input type='text' size='20' id='new_downlFrom' name='new_downlFrom' value='原创'></td>
            <td align="right">来自 URL: </td>  
            <td><input type='text' size='20' id='new_downlFromURL' name='new_downlFromURL' value='#'></td>
          </tr>

          <tr class='formbox-content'> 
            <td align="right">作者:</td>
            <td><input type='text' size='20' id='new_downlAuthor' name='new_downlAuthor' value='<%=memName%>'></td>
            <td align="right">Image URL:</td>
            <td colspan="2"><input type='text' size='20' id='new_downlImgPath' name='new_downlImgPath'></td>
          </tr>
          <tr class='formbox-content'> 
            <td height="31" align="right">地址1说明:</td>
            <td><input type='text' size='20' id='new_downlDcomm1' name='new_downlDcomm1'><font color="#FF0000">*</font></td>
            <td align="right"><font color="#FF0000">*</font>地址1:</td>
            <td colspan="2"><input type='text' size='20' maxlength="250" id='new_downlDcomm1URL' name='new_downlDcomm1URL'></td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">地址2说明:</td>
            <td><input type='text' size='20' id='new_downlDcomm2' name='new_downlDcomm2'></td>
            <td align="right">地址2:</td>
            <td colspan="2"><input type='text' size='20' maxlength="250" id='new_downlDcomm2URL' name='new_downlDcomm2URL'></td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">地址3说明:</td>
            <td><input type='text' size='20' id='new_downlDcomm3' name='new_downlDcomm3'></td>
            <td align="right">地址3:</td>
            <td colspan="2"><input type='text' size='20' maxlength="250" id='new_downlDcomm3URL' name='new_downlDcomm3URL'></td>
          </tr>
          <tr class='formbox-content'> 
            <td align="right">地址4说明:</td>
            <td><input type='text' size='20' id='new_downlDcomm4' name='new_downlDcomm4'></td>
            <td align="right">地址4:</td>
            <td colspan="2"><input type='text' size='20' maxlength="250" id='new_downlDcomm4URL' name='new_downlDcomm4URL'></td>
          </tr>
          <tr class='formbox-content'> 
            <td height="53" align="right" valign="top">路径(地址):</td>
            <td colspan="4" align="left" valign="middle"> <input type='text' size='80' id='downl_site' name='downl_site'> 
              <br>
              通过上传功能添加图片和文件。然后复制地址到所需文本框。 支持文件格式:<font color="#990000">gif,jpg,bmp,png,tif,rar,zip</font></td>
          </tr>
         
          <tr class='formbox-content'> 
            <td align="right">上传:</td>
            <td colspan="4"><iframe border="0" frameBorder="0" frameSpacing="0" height="21" marginHeight="0" marginWidth="0" noResize scrolling="no" width="100%" vspale="0" src="attachment.asp"></iframe></td>
          </tr>
        
          <tr class='formbox-content'> 
            <td height="59" align="right" valign="top">介绍:</td>
            <td colspan="4" align="left" valign="top"> <textarea name="new_downlComment" cols="80" rows="6" wrap="VIRTUAL" id="new_downlComment" value="&nbsp;"></textarea></td>
          </tr>
          <tr class='formbox-content'> 
            <td colspan="5" align="center"><input type="Submit" name="Submit" value=" 保存更改 "></td>
          </tr>
        </form>
        <%End IF%>
</table>

<%ElseIF Request.QueryString("action")="tag" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center"><form name="tag_Search" method="get" action="admincp.asp">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> TAG管理</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> Tag 管理&nbsp;&nbsp;|&nbsp;&nbsp;<input type="text" name="sTXT"><input type="hidden" name="action" value="tag">&nbsp;<input type="submit" value="确定">&nbsp;&nbsp;|&nbsp;&nbsp;<%If Request.QueryString("type")="count" Then
	Call CounTag
	Response.Write("<a href=""admincp.asp?action=tag"">整理 TAGS 成功，请点击返回</a>")
	Else
	Response.Write("<a href=""admincp.asp?action=tag&type=count"">整理 TAGS</a>")
	End If%>	</td>
  </tr></form>
  <tr>
    <td bgcolor="#FFFFFF"><%If Request.QueryString("type")="tagedit" Then
		If CheckStr(Request.QueryString("tagID"))=Empty OR CheckStr(Request.Form("edit_tagsName"))=Empty OR CheckStr(Request.Form("edit_blog_ID"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">请将信息填写完整，点击返回上一页</a>")
		Else
			If CheckStr(Request.Form("edit_tagDele"))="True" Then
				Conn.ExeCute("DELETE  FROM blog_tag WHERE TagsID="&CheckStr(Request.QueryString("tagID"))&"")
				SQLQueryNums=SQLQueryNums+1
				'删除或减少
				Dim tagIfedit
				Set tagIfedit=Server.CreateObject("Adodb.Recordset")
				SQL="SELECT TagName,TagBlogCount FROM blog_Tags Where TagName='"&CheckStr(Request.Form("edit_tagsName"))&"'"
				tagIfedit.Open SQL,CONN,1,1
				SQLQueryNums=SQLQueryNums+1
				If tagIfedit.EOF AND tagIfedit.BOF Then
				Else
				  If tagIfedit("TagBlogCount")-1=0 Then
					Conn.ExeCute("DELETE  FROM blog_tags WHERE TagName='"&CheckStr(Request.Form("edit_tagsName"))&"'")
					SQLQueryNums=SQLQueryNums+1
			 	  ElseIf tagIfedit("TagBlogCount")-1>0 Then
					Conn.ExeCute("UPDATE blog_Tags SET TagBlogCount=TagBlogCount-1 WHERE TagName='"&CheckStr(Request.Form("edit_tagsName"))&"'")
					SQLQueryNums=SQLQueryNums+1
			  	  End If
			  	End If
			  	tagIfedit.Close
			  	Set tagIfedit=Nothing
			  	'删除或减少 End
			Else
				Conn.ExeCute("UPDATE blog_Tag SET TagsName='"&CheckStr(Request.Form("edit_tagsName"))&"',blog_ID='"&CheckStr(Request.Form("edit_blog_ID"))&"' WHERE TagsID="&CheckStr(Request.QueryString("tagID"))&"")
				SQLQueryNums=SQLQueryNums+1
				'添加或整理
				Dim tagIfchan
				Set tagIfchan=Server.CreateObject("Adodb.Recordset")
				SQL="SELECT TagName,TagBlogCount FROM blog_Tags Where TagName='"&CheckStr(Request.Form("edit_tagsName"))&"'"
				tagIfchan.Open SQL,CONN,1,1
				SQLQueryNums=SQLQueryNums+1
				If tagIfchan.EOF AND tagIfchan.BOF Then
					Conn.ExeCute("INSERT INTO blog_Tags(TagName) VALUES ('"&CheckStr(Request.Form("edit_tagsName"))&"')")
					SQLQueryNums=SQLQueryNums+1
				End If
				Call CounTag
				tagIfchan.Close
				Set tagIfchan=Nothing
				'添加或整理 End
			End If
			Response.Write("<a href=""admincp.asp?action=tag"">编辑 TAG 成功，请点击返回</a>")
		End If
	ElseIF Request.QueryString("type")="tagadd" Then
		If CheckStr(Request.Form("new_tagsName"))=Empty OR CheckStr(Request.Form("new_blog_ID"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">请将信息填写完整，点击返回上一页</a>")
		Else
			'添加或增加
			Dim log_tags
			log_tags=Split(CheckStr(Request.Form("new_tagsName")),";")
			Dim tabs,post_tags
			For tabs=0 To Ubound(log_tags)
			    SQL="SELECT TagName From blog_Tags WHERE TagName='"&log_tags(tabs)&"'"
				Set post_tags=Server.CreateObject("Adodb.Recordset")
				post_tags.OPEN SQL,Conn,1,1
				SQLQueryNums=SQLQueryNums+1
				If post_tags.EOF And post_tags.BOF Then
				   Conn.ExeCute("INSERT Into blog_Tags(TagName) Values ('"&log_tags(tabs)&"')")
				   SQLQueryNums=SQLQueryNums+1
				Else
				   Conn.ExeCute("UPDATE blog_Tags Set TagBlogCount=TagBlogCount+1 WHERE TagName='"&log_tags(tabs)&"'")
				   SQLQueryNums=SQLQueryNums+1
				End If
				Conn.ExeCute("INSERT Into blog_Tag(blog_ID,TagsName) Values("&CheckStr(Request.Form("new_blog_ID"))&",'"&log_tags(tabs)&"')")
				SQLQueryNums=SQLQueryNums+1
				post_tags.Close
			    Set post_tags=Nothing
			Next
			'添加或增加 End
			Response.Write("<a href=""admincp.asp?action=tag"">添加 TAG 成功，请点击返回</a>")
		End If
	Else
		Dim blog_tag,blog_tagNums,blog_tagSQL,blog_tagURL
		If Search_TXT<>Empty Then
			blog_tagSQL=" And (TagsName LIKE '%"&Search_TXT&"%' OR blog_ID LIKE '%"&Search_TXT&"%')"
			blog_tagURL="sTXT="&Search_TXT&"&"
		End If
		Set blog_tag=Server.CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM blog_Tag WHERE TagsID<>null"&blog_tagSQL&" ORDER BY TagsID DESC"
		blog_tag.Open SQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		If blog_tag.EOF And blog_tag.BOF Then
			Response.Write("暂时没有 TAG")
		Else

			IF CheckStr(Request.QueryString("Page"))<>Empty Then
				Curpage=Cint(CheckStr(Request.QueryString("Page")))
				IF Curpage<0 Then Curpage=1
			Else
				Curpage=1
			End IF

			Url_Add="?action=tag&"&blog_tagURL
			blog_tag.PageSize=15
			blog_tag.AbsolutePage=CurPage
			blog_tagNums=blog_tag.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(blog_tagNums,15,CurPage,Url_Add)&"</span>"
			Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr bgcolor=""#EFEFEF""><td>删？</td><td>名称</td><td>日志</td><td width=""100%"">操作</td></tr>")
			Do Until blog_tag.EOF Or PageCount=15
				Response.Write("<tr bgcolor=""#FFFFFF""><form name=""addtag"" action=""admincp.asp?action=tag&type=tagedit&tagID="&blog_tag("TagsID")&""" method=""post""><td><input name=""edit_tagDele"" type=""checkbox"" id=""edit_tagDele"" value=""True""></td><td><input type=""text"" size=""24"" id=""edit_tagsName"" name=""edit_tagsName"" value="""&blog_tag("TagsName")&"""><input type=""hidden"" id=""edit_tagsNameOld"" name=""edit_tagsNameOld"" value="""&blog_tag("TagsName")&"""></td><td nowrap><input type=""text"" size=""8"" id=""edit_blog_ID"" name=""edit_blog_ID"" value="""&blog_tag("blog_ID")&""">&nbsp;&nbsp;<a href=""blogview.asp?logID="&blog_tag("blog_ID")&""" target=""_blank"">查看所在日志</a>&nbsp;</td><td><input type=""Submit"" name=""Submit"" value="" 编 辑 "" class=""button""></td></form></tr>")
				blog_tag.MoveNext
				PageCount=PageCount+1
			Loop			
			Response.Write("</table>")
		Response.Write(MultiPages)
		End If
		blog_tag.Close
		Set blog_tag=Nothing
		Response.Write("<hr width=""100%"" size=""1""><table><tr><form name=""addtag"" action=""admincp.asp?action=tag&type=tagadd"" method=""post""><td>名称：<input type=""text"" size=""28"" id=""new_tagsName"" name=""new_tagsName""> 最多5个 Tag 用;分隔&nbsp;&nbsp;日志：<input type=""text"" size=""8"" id=""new_blog_ID"" name=""new_blog_ID"">&nbsp;&nbsp;<input type=""Submit"" name=""Submit"" value="" 添  加 "" class=""button""></td></form></tr></table>")
	End If%></td>
  </tr>
</table>

<%ElseIF Request.QueryString("action")="tags" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center"><form name="tag_Search" method="get" action="admincp.asp">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> Tags 管理&nbsp;&nbsp;|&nbsp;&nbsp;<input type="text" name="sTXT"><input type="hidden" name="action" value="tags">&nbsp;<input type="submit" value="确定">&nbsp;&nbsp;|&nbsp;&nbsp;<%If Request.QueryString("type")="count" Then
	Call CounTag
	Response.Write("<a href=""admincp.asp?action=tags"">整理 TAGS 成功，请点击返回</a>")
	Else
	Response.Write("<a href=""admincp.asp?action=tags&type=count"">整理 TAGS</a>")
	End If%>	
	</td>
  </tr></form>
  <tr>
    <td bgcolor="#FFFFFF"><%If Request.QueryString("type")="tagsedit" Then
		If CheckStr(Request.QueryString("TagID"))=Empty OR CheckStr(Request.Form("edit_tagName"))=Empty Then
			Response.Write("<a href=""javascript:history.go(-1);"">请将信息填写完整，点击返回上一页</a>")
		Else
			If CheckStr(Request.Form("edit_TagDele"))="True" Then
				Conn.ExeCute("DELETE  FROM blog_Tags WHERE TagID="&CheckStr(Request.QueryString("TagID"))&"")
				SQLQueryNums=SQLQueryNums+1
				Conn.ExeCute("DELETE  FROM blog_Tag WHERE TagsName='"&CheckStr(Request.Form("edit_tagName"))&"'")
				SQLQueryNums=SQLQueryNums+1
			Else
				Conn.ExeCute("UPDATE blog_Tags SET TagName='"&CheckStr(Request.Form("edit_tagName"))&"' WHERE TagName='"&CheckStr(Request.Form("edit_tagNameOld"))&"'")
				SQLQueryNums=SQLQueryNums+1
				Conn.ExeCute("UPDATE blog_Tag SET TagsName='"&CheckStr(Request.Form("edit_tagName"))&"' WHERE TagsName='"&CheckStr(Request.Form("edit_tagNameOld"))&"'")
				SQLQueryNums=SQLQueryNums+1
			End If
			Response.Write("<a href=""admincp.asp?action=tags"">编辑 TAGS 成功，请点击返回</a>")
		End If
	Else
		Dim blog_tags,blog_tagsNums,blog_tagsSQL,blog_tagsURL
		If Search_TXT<>Empty Then
			blog_tagsSQL=" And (TagName LIKE '%"&Search_TXT&"%')"
			blog_tagsURL="sTXT="&Search_TXT&"&"
		End If
		Set blog_tags=Server.CreateObject("ADODB.RecordSet")
		SQL = "SELECT * FROM blog_Tags WHERE TagID<>null"&blog_tagsSQL&" ORDER BY TagID DESC"
		blog_tags.Open SQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		If blog_tags.EOF And blog_tags.BOF Then
			Response.Write("暂时没有 TAGS")
		Else
			IF CheckStr(Request.QueryString("Page"))<>Empty Then
				Curpage=Cint(CheckStr(Request.QueryString("Page")))
				IF Curpage<0 Then Curpage=1
			Else
				Curpage=1
			End IF
			Url_Add="?action=tags&"&blog_tagsURL
			blog_tags.PageSize=15
			blog_tags.AbsolutePage=Curpage
			blog_tagsNums=blog_tags.RecordCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(blog_tagsNums,15,Curpage,Url_Add)&"</span>"
			Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr bgcolor=""#EFEFEF""><td>删？</td><td>名称</td><td nowrap>数量</td><td>创建时间</td><td width=""100%"">操作</td></tr>")
			Do Until blog_tags.EOF Or PageCount=15
				Response.Write("<tr bgcolor=""#FFFFFF""><form name=""addtag"" action=""admincp.asp?action=tags&type=tagsedit&tagID="&blog_tags("TagID")&""" method=""post""><td><input name=""edit_tagDele"" type=""checkbox"" id=""edit_tagDele"" value=""True""></td><td><input type=""text"" size=""24"" id=""edit_tagName"" name=""edit_tagName"" value="""&blog_tags("TagName")&"""><input type=""hidden"" id=""edit_tagNameOld"" name=""edit_tagNameOld"" value="""&blog_tags("TagName")&"""></td><td nowrap>"&blog_tags("TagBlogCount")&"</td><td nowrap><span class=""date"">"&blog_tags("CreateDate")&"</span></td><td><input type=""Submit"" name=""Submit"" value="" 编 辑 "" class=""button""></td></form></tr>")
				blog_tags.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("</table>")
			Response.Write(MultiPages)
		End If
		blog_tags.Close
		Set blog_tags=Nothing
	End If%></td>
  </tr>
</table>

<%ElseIF Request.QueryString("action")="member" Then%><br>
<table width="99%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td bgcolor="#FFFFFF" class="siderbar_head"><%=SiteName%> 会员管理</td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><%IF Request.QueryString("type")="editmem" Then
		IF CheckStr(Request.QueryString("memID"))=Empty AND CheckStr(Request.QueryString("memType"))<>"Member" AND CheckStr(Request.QueryString("memType"))<>"Admin" AND CheckStr(Request.QueryString("memType"))<>"SupAdmin" Then
			Response.Write("<a href=""admincp.asp?action=member"">对不起，参数错误，点击返回会员管理</a>")
		Else
			Conn.ExeCute("UPDATE blog_Member SET mem_Status='"&CheckStr(Request.QueryString("memType"))&"' WHERE mem_ID="&CheckStr(Request.QueryString("memID")))
			SQLQueryNums=SQLQueryNums+1
			Response.Write("<a href=""admincp.asp?action=member"">编辑会员成功，点击返回会员管理</a>")
		End IF
	ElseIF Request.QueryString("type")="delemem" Then
		IF CheckStr(Request.QueryString("memID"))=Empty Then
			Response.Write("<a href=""admincp.asp?action=member"">对不起，参数错误，点击返回会员管理</a>")
		Else
			Conn.ExeCute("DELETE  FROM blog_Member WHERE mem_ID="&CheckStr(Request.QueryString("memID")))
			Conn.ExeCute("UPDATE blog_Info SET blog_MemNums=blog_MemNums-1")
			SQLQueryNums=SQLQueryNums+2
			Response.Write("<a href=""admincp.asp?action=member"">删除会员成功，点击返回会员管理</a>")
		End IF
	Else
		Dim adm_MemList
		Set adm_MemList=Server.CreateObject("Adodb.RecordSet")
		SQL="SELECT mem_ID,mem_Name,mem_Regtime,mem_Status FROM blog_Member ORDER BY mem_Regtime DESC"
		adm_MemList.Open SQL,Conn,1,1
		SQLQueryNums=SQLQueryNums+1
		IF adm_MemList.EOF AND adm_MemList.BOF Then
			Response.Write("暂时没有会员")
		Else
			Dim CurPage,Url_Add
			IF CheckStr(Request.QueryString("Page"))<>Empty Then
				Curpage=Cint(CheckStr(Request.QueryString("Page")))
				IF Curpage<0 Then Curpage=1
			Else
				Curpage=1
			End IF
			Url_Add="?action=member&"
			adm_MemList.PageSize=15
			adm_MemList.AbsolutePage=CurPage
			Dim adm_MemNums
			adm_MemNums=adm_MemList.RecordCount
			Dim MultiPages,PageCount
			MultiPages="<span class=""smalltxt"">"&MultiPage(adm_MemNums,15,CurPage,Url_Add)&"</span>"
			Response.Write(MultiPages)
			Response.Write("<table width=""100%"" border=""0"" align=""center"" cellpadding=""4"" cellspacing=""1"" bgcolor=""#CCCCCC""><tr bgcolor=""#EFEFEF""><td nowrap>编号</td><td nowrap>会员名称</td><td nowrap>会员身份</td><td nowrap>注册时间</td><td width=""100%"">会员操作</td></tr>")
			Do Until adm_MemList.EOF OR PageCount=15
				Dim adm_MemStatus,adm_MemStatusAct
				Select Case adm_MemList("mem_Status")
				Case "SupAdmin"
					adm_MemStatus="超级管理员"
					adm_MemStatusAct="<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=Member"">一般会员</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=Admin"">设为一般管理员</a>"
				Case "Admin"
					adm_MemStatus="一般管理员"
					adm_MemStatusAct="<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=Member"">一般会员</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=SupAdmin"">设为超级管理员</a>"
				Case "Member"
					adm_MemStatus="一般会员"
					adm_MemStatusAct="<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=Admin"">一般管理员</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=""admincp.asp?action=member&type=editmem&memID="&adm_MemList("mem_ID")&"&memType=SupAdmin"">设为超级管理员</a>"
				End Select
				Response.Write("<tr bgcolor=""#FFFFFF""><td nowrap>"&adm_MemList("mem_ID")&"</td><td nowrap><a href=""member.asp?action=view&memName="&Server.URLEncode(adm_MemList("mem_Name"))&""" target=""_blank"" alt=""点击查看会员资料"">"&adm_MemList("mem_Name")&"</td><td nowrap>"&adm_MemStatus&"</td><td nowrap>"&DateToStr(adm_MemList("mem_RegTime"),"Y-m-d H:I A")&"</td><td>&nbsp;<a href=""admincp.asp?action=member&type=delemem&memID="&adm_MemList("mem_ID")&""">删除会员</a>&nbsp;&nbsp;|&nbsp;&nbsp;"&adm_MemStatusAct&"</td></tr>")
				adm_MemList.MoveNext
				PageCount=PageCount+1
			Loop
			Response.Write("</table>")
			Response.Write(MultiPages)
		End IF
		adm_MemList.Close
		Set adm_MemList=Nothing
	End IF%></td>
  </tr>
</table>
<%Else%><br>
<br>
<table width="92%" border="0" align="center" cellpadding="6" cellspacing="1" bgcolor="#CCCCCC" align="center">
  <tr>
    <td colspan="2" class="msg_head"><%=SiteName%> 服务器基本信息</td>
</tr>
  <tr>
    <td bgcolor="#FFFFFF" width="40%" nowrap>&nbsp;服务器时间：<%=DateToStr(Now(),"Y-m-d H:I A")%></td>
    <td bgcolor="#FFFFFF" width="60%">&nbsp;服务器IP地址：<%=Request.ServerVariables("LOCAL_ADDR")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;服务器IIS版本：<%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
    <td bgcolor="#FFFFFF">&nbsp;脚本解释引：<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;服务器空间占用：<%=GetTotalSize(Request.ServerVariables("APPL_PHYSICAL_PATH"),"Folder")%></td>
    <td bgcolor="#FFFFFF">&nbsp;站点物理路径：<%=Request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;服务器CPU数量：<%=GetSysInfo("CPUNums")%></td>
    <td bgcolor="#FFFFFF">&nbsp;脚本超时设置：<%=Server.ScriptTimeout%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF">&nbsp;服务器操作系统：<%=GetSysInfo("OSInfo")%></td>
</tr>
  <tr>
    <td colspan="2" bgcolor="#FFFFFF">&nbsp;服务器CPU信息：<%=GetSysInfo("CPUInfo")%></td>
</tr>
  <tr>
    <td colspan="2" class="msg_head">服务器组件安装情况</td>
</tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;FileUp上传组件：<%=CheckObjInstalled("FileUp.upload")%></td>
    <td bgcolor="#FFFFFF">&nbsp;FSO文本读写：<%=CheckObjInstalled("Scripting.FileSystemObject")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;数据库使用：<%=CheckObjInstalled("adodb.connection")%></td>
    <td bgcolor="#FFFFFF">&nbsp;Jmail组件支持：<%=CheckObjInstalled("JMail.SMTPMail")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;GflSDK组件支持：<%=CheckObjInstalled("GflAx190.GflAx")%></td>
    <td bgcolor="#FFFFFF">&nbsp;EasyMail邮件支持：<%=CheckObjInstalled("easymail.Mailsend")%></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF">&nbsp;无组件上传－ADODB.Stream：<%=CheckObjInstalled("Scripting.Dictionary")%></td>
    <td bgcolor="#FFFFFF">&nbsp;无组件上传－Scripting.Dictionary ：<%=CheckObjInstalled("Scripting.Dictionary")%></td>
  </tr>
</table>
<%
Function GetSysInfo(InfoType)
  Dim WshShell,WshSysEnv
  Set WshShell = Server.CreateObject("WScript.Shell")
  Set WshSysEnv = WshShell.Environment("SYSTEM")
  If InfoType="CPUNums" Then
		GetSysInfo=Cstr(WshSysEnv("NUMBER_OF_PROCESSORS"))
		If IsNull(GetSysInfo) Then
			GetSysInfo = Request.ServerVariables("NUMBER_OF_PROCESSORS")
		ElseIf GetSysInfo="" Then
			GetSysInfo = Request.ServerVariables("NUMBER_OF_PROCESSORS")
		End If
	ElseIf InfoType="CPUInfo" Then
		GetSysInfo = Cstr(WshSysEnv("PROCESSOR_IDENTIFIER"))
	ElseIf InfoType="OSInfo" Then
		GetSysInfo = Cstr(WshSysEnv("OS"))
		If Request.ServerVariables("OS")="" Then GetSysInfo=GetSysInfo & "(可能是 Windows Server 2003)"
	End If 
End Function

Function CheckObjInstalled(strClassString)
	On Error Resume Next
	Dim TmpObj
	Set TmpObj = Server.CreateObject(strClassString)
	IF Err = 0 OR Err = -2147221477 Then
		CheckObjInstalled= "<font color=""#00FF00""><b>√</b></font>"
	ElseIF Err = 1 OR Err = -2147221005 Then
		CheckObjInstalled="<font color=""#FF0000""><b>×</b></font>"
	End IF
	Err.Clear
	Set TmpObj = Nothing
End Function

Function GetTotalSize(GetLocal,GetType)
	Dim FSO
	Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		GetTotalSize="服务器关闭FSO，查看占用空间失败"
	Else
		Dim SiteFolder
		IF GetType="Folder" Then
			Set SiteFolder=FSO.GetFolder(GetLocal) 
		Else
			Set SiteFolder=FSO.GetFile(GetLocal) 
		End IF
		GetTotalSize=SiteFolder.Size
		IF GetTotalSize>1024*1024 Then
		GetTotalSize=GetTotalSize/1024/1024
		IF inStr(GetTotalSize,".") Then GetTotalSize = Left(GetTotalSize,inStr(GetTotalSize,".")+2)
			GetTotalSize=GetTotalSize&" MB"
		Else
			GetTotalSize=Fix(GetTotalSize/1024)&" KB"
		End IF
		Set SiteFolder=Nothing
	End IF
	Set FSO=Nothing
End Function

End IF
End IF
Function DeleteFiles(FilePath)
    Dim FSO
    Set FSO=Server.CreateObject("Scripting.FileSystemObject")
	IF Err<>0 Then
		Err.Clear
		Response.Write("服务器关闭FSO，无法删除文件")
	Else
		If FSO.FileExists(FilePath) Then
			FSO.DeleteFile FilePath,True
			DeleteFiles = 1
		Else
			DeleteFiles = 0
		End If
	End If
    Set FSO = Nothing
End Function
Function FreeApplicationMemory
	Response.Write "<b>释放网站数据列表：</b>" & VbCrLf
	Dim Thing
	For Each Thing IN Application.Contents
		IF Left(Thing,Len(CookieName)) = CookieName Then
			Response.Write "<font color=""gray"">" & thing & "</font><br>"
			IF isObject(Application.Contents(Thing)) Then
				Application.Contents(Thing).Close
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = Null
				Response.Write "成功关闭对象"
			ElseIF isArray(Application.Contents(Thing)) Then
				Set Application.Contents(Thing) = Nothing
				Application.Contents(Thing) = Null
				Response.Write "成功释放数组"
			Else
				Response.Write(HtmlEncode(Application.Contents(Thing)))
				Application.Contents(Thing) = Null
			End IF
			Response.Write("&nbsp;&nbsp;")
		End IF
	Next
End Function%></td>
  </tr>
</table><%End IF%></td>
  </tr>
</table>
<!--#include file="footer.asp" -->