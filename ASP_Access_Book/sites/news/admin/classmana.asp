<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="conn.asp"-->
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>您没有操作的权限</p>"
	else
		call main()
	end if
  end if
dim sql,rs
dim layer,child,rootID,parentID,Maxrootid%>
<%sub main()%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>类别管理</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body>
<div align="center">
  <center>
<table border="0" cellPadding="3" cellSpacing="0" class="border" width="95%" style="border-collapse: collapse" bordercolor="#111111" >
  <tr>
    <th class="title" colSpan="2" height="25" width="100%">类别管理 </th>
  </tr>
  <tr>
    <td class="tdbg" colSpan="2"><b>注意</b>：<br>
    ①删除类别同时将删除该类别下所有资料！删除分类同时删除下属分类和其中资料！ 操作时请完整填写表单信息。<br>
    ②如果选择<b>复位所有类别</b>，则所有类别都将作为一级类别分类，这时您需要重新对各个类别进行归属的基本设置，<b>不要轻易使用该功能</b>，仅在做出了错误的设置而无法复原类别之间的关系和排序的时候使用 
    </td>
  </tr>
  <tr>
    <td class="title"><b>类别操作选项</b></td>
    <td class="title"><a href="classmana.asp">
    类别管理首页</a> | <a href="?action=add">
    新建一级类别</a> | </td>
  </tr>
</table>
<p></p>
<%
select case Request("action")
case "add"
	call add()
case "edit"
	call edit()
case "savenew"
	call savenew()
case "savedit"
	call savedit()
case "orders"
	call orders()

case "delclass"
	call delclass()
case "updatclassorders"
	call updateclassorders()

case "del"
	call del()
case else
	call classinfo()
end select
end sub

sub classinfo()
sql="select * from class order by rootID,ordersID"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
rs.close
call add()
else%>
<table width="95%" cellspacing="0" cellpadding="0" class="border" style="border-collapse: collapse" bordercolor="#111111">
<tr> 
<th width="39%" class="title" height=25>类别分类
</th>
<th width="31%" class="title" height=25>操作
</th>
</tr>
<%j=0
do while not rs.eof
%>
<%if j mod 2 = 1 then%><tr class="tdbg"><%else%><tr class="tdbg2"><%end if %>
<td height="25" width="39%">
<%if rs("layer")>0 then%>
<%for i=1 to rs("layer")%>
&nbsp;
<%next%>
<%end if%>
<%if rs("child")>0 then%><img src="pic/plus.gif"><%else%><img src="pic/nofollow.gif"><%end if%>
<%if rs("parentID")=0 then%><b><%end if%><%=rs("class")%><%if rs("child")>0 then%>(<%=rs("child")%>)<%end if%>
</td>
<td width="61%">
<p align="left"><a href="classmana.asp?action=add&editid=<%=rs("classid")%>"><U>添加分类</U></a> | <a href="classmana.asp?action=edit&editid=<%=rs("classid")%>"><U>修改类别</U></a> 
<%if rs("child")>0 then%>| <a href="classmana.asp?action=orders&editid=<%=rs("classid")%>"><U>排序</U></a> <%end if%>| <%if rs("child")=0 then%><a href="classmana.asp?action=delclass&classid=<%=rs("classid")%>" onclick="{if(confirm('清空将包括该类别的所有资料，确定清空吗?')){return true;}return false;}"><U>清空</U></a> | <a href="classmana.asp?action=del&editid=<%=rs("classid")%>" onclick="{if(confirm('删除将包括该类别的所有资料，确定删除吗?')){return true;}return false;}"><U>删除</U></a>&nbsp;<%end if%></td>
</tr>
<%
rs.movenext
j=j+1
loop
end if
%>
</table>
<%
'rs.close
end sub%>
<%sub add()
dim classnum
	sql="select Max(classid) from class"
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
	classnum=1
	else
	classnum=rs(0)+1
	end if
	if isnull(classnum) then classnum=1 end if
	rs.close
%>
<form action="classmana.asp?action=savenew" method="post">
<input type="hidden" name="newclassid" value=<%=classnum%>>
<table width="95%" border="0" cellspacing="0" cellpadding="3" class="border" style="border-collapse: collapse" bordercolor="#111111">
<tr> 
<th height=24 colspan=2 class="title"><B>添加类别</th>
</tr>
<tr> 
<td width="52%" height=30 class="tdbg" align="center">类别名称</td>
<td width="48%" class="tdbg"> 
<input type="text" name="class" size="35" class="smallinput">
</td>
</tr>
<tr> 
<td width="52%" height=30 class="tdbg" align="center"><U>所属类别</U></td>
<td width="48%" class="tdbg"> 
<select name=ParentID class="select">
<option value="0">没有父类</option>
<%sql="select * from class order by parentID,ordersID"
set rs=conn.execute(sql)
do while not rs.EOF%>
<option value="<%=rs("classid")%>" <%if request("editid")<>"" and clng(request("editid"))=rs("classid") then%>selected<%end if%>>
<%if rs("layer")>0 then%>
<%for i=1 to rs("layer")%>
－
<%next%>
<%end if%><%=rs("class")%></option>
<%
rs.MoveNext 
loop
rs.Close 
%>
</select> <select size="1" name="classtype" class="select">
<option value="0">栏目</option>
<option value="1">文章</option>
<option value="2">学者/专家</option>
<option value="3">网站连接</option>
</select></td>
</tr>
<tr> 
<td width="100%" height=24 class="tdbg" colspan="2" align="center"> 
<input type="submit" name="Submit" value="添加类别" class="buttonface">
</td>
</tr>
</table></form>
<%end sub
sub edit()
sql = "select class,parentID,classtype from class where classid="&request("editid")
set rs=conn.execute(sql)
dim classname,parentID
classname=rs(0)
parentID=rs(1)
classtype=rs(2)
rs.close%>
<form action="classmana.asp?action=savedit" method="post">
<table width="95%" border="0" cellspacing="0" cellpadding="3" class="border" style="border-collapse: collapse" bordercolor="#111111">
<tr> 
<th height=24 colspan=2 class="title"><B>修改类别 <%=classname%>
<input type="hidden" name=editid value="<%=Request("editid")%>"></th>
</tr>
<tr> 
<td width="52%" height=30 class="tdbg" align="center">类别名称</td>
<td width="48%" class="tdbg"> 
<input type="text" name="class" size="35" class="smallinput" value="<%=classname%>">
</td>
</tr>
<tr> 
<td width="52%" height=30 class="tdbg" align="center"><U>所属类别</U></td>
<td width="48%" class="tdbg"> 
<select name=ParentID class="select">
<option value="0">没有父类</option>
<%sql="select * from class order by rootID,ordersID,layer"
set rs=conn.execute(sql)
do while not rs.EOF%>
<option value="<%=rs("classid")%>" <%if parentID<>"" and clng(parentID)=rs("classid") then%>selected<%end if%>>
<%if rs("layer")>0 then%>
<%for i=1 to rs("layer")%>
－
<%next%>
<%end if%><%=rs("class")%></option>
<%
rs.MoveNext 
loop
rs.Close 
%>
</select>
<select size="1" name="classtype" class="select">
<option value="0"<%if classtype=0 then response.write " selected" end if%>>栏目</option>
<option value="1"<%if classtype=1 then response.write " selected" end if%>>文章</option>
<option value="2"<%if classtype=2 then response.write " selected" end if%>>学者/专家</option>
<option value="3"<%if classtype=3 then response.write " selected" end if%>>网站链接</option>
</select></td>
</tr>
<tr> 
<td width="100%" height=24 class="tdbg" colspan="2" align="center"> 
<input type="hidden" name="oldparentID" value='<%=parentID%>'>
<input type="submit" name="Submit" value="修改类别" class="buttonface">
</td>
</tr>
</table></form>
<%end sub
sub orders()
%>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="title">
<tr> 
	<th height="22">分类重新排序修改
</th>
</tr>
	<tr>
	<td class="tdbg"><table width="90%">
<%sql="select * from class where parentID="&request("editid")&" order by RootID,ordersID"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	response.write "还没有相应的分类。"
else
do while not rs.eof
	response.write "<form action=classmana.asp?action=updatclassorders method=post><tr><td width=""50%""><input type=hidden name=parentID value="&request("editID")&">"
	if rs("layer")>0 then
	for i=1 to rs("layer")
		response.write "&nbsp;"
	next
	end if
	if rs("child")>0 then
		response.write "<img src=pic/plus.gif>"
	else
		response.write "<img src=pic/nofollow.gif>"
	end if
	if rs("parentid")=0 then
		response.write "<b>"
	end if
	response.write rs("class")
	if rs("child")>0 then
		response.write "("&rs("child")&")"
	end if
	response.write "</td><td width=""50%"">"
	if rs("ParentID")>0 then
	set trs=conn.execute("select count(*) from class where ParentID="&request("editid")&" and ordersID<"&rs("ordersID")&"")
	uporders=trs(0)
	if isnull(uporders) then uporders=0
	set trs=conn.execute("select count(*) from class where ParentID="&request("editid")&" and ordersID>"&rs("ordersID")&"")
	doorders=trs(0)
	if isnull(doorders) then doorders=0
	if uporders>0 then
		response.write "<select name=uporders size=1><option value=0>向上移动</option>"
		for i=1 to uporders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 then
		if uporders>0 then response.write "&nbsp;"
		response.write "<select name=doorders size=1><option value=0>向下移动</option>"
		for i=1 to doorders
		response.write "<option value="&i&">"&i&"</option>"
		next
		response.write "</select>"
	end if
	if doorders>0 or uporders>0 then
	response.write "<input type=hidden name=""editID"" value="""&rs("classid")&""">&nbsp;<input type=submit name=Submit value=修改>"
	end if
	end if
	response.write "</td></tr></form>"
	uporders=0
	doorders=0
	rs.movenext
	loop
	response.write "</table>"
end if
rs.close
set rs=nothing
%>
</td></tr></table>
<%
end sub
sub savenew()
if request("class")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入类别名称。"
	Founderr=true
end if
if request("parentID")="" then
	Errmsg=Errmsg+"<br>"+"<li>请选择父类或选择“没有父类”。"
	Founderr=true
end if
if founderr=true then
	response.write Errmsg
	exit sub
end if
if request("parentID")<>"0" then
set rs=conn.execute("select rootid,classid,layer,ordersID,ParentStr from class where classid="&request("parentID"))
rootid=rs(0)
parentid=rs(1)
layer=rs(2)
ordersID=rs(3)
if layer+1>20 then
	response.write "本系统限制最多只能有20级分类"
	exit sub
end if
parentstr=rs(4)
else
set rs=conn.execute("select max(rootid) from class")
maxrootid=rs(0)+1
if isnull(MaxRootID) then
 MaxRootID=1
end if
end if
sql="select classid from class where classid="&request("newclassid")
set rs=conn.execute(sql)
if not (rs.eof and rs.bof) then
	response.write "您不能指定和别的类别一样的序号。"
	exit sub
else
	classid=request("newclassid")
end if
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from class"
rs.Open sql,conn,1,3
rs.AddNew
if request("parentID")<>"0" then
	rs("layer")=layer+1
	rs("rootid")=rootid
	rs("ordersID") = Request.form("newclassid")
	rs("parentid") = Request.Form("parentID")
	if ParentStr="0" then
		rs("ParentStr")=Request.Form("parentID")
	else
		rs("ParentStr")=ParentStr & "," & Request.Form("parentID")
	end if
else
	rs("layer")=0
	rs("rootid")=maxrootid
	rs("ordersID")=0
	rs("parentid")=0
	rs("child")=0
	rs("parentstr")=0
end if
rs("class")=request("class")
rs("classid") = Request.form("newclassid")
rs("classtype") = Request.form("classtype")
rs.Update 
rs.Close
if request("parentID")<>"0" then
if layer>0 then
	'当上级分类深度大于0的时候要更新其父类（或父类的父类）的版面数和相关排序
	for i=1 to layer
		'更新其父类版面数
		if parentid<>"" then
		conn.execute("update class set child=child+1 where classid="&parentid)
		end if
		'得到其父类的父类的版面ID
		set rs=conn.execute("select parentid from class where classid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
		end if
		'当循环次数大于1并且运行到最后一次循环的时候直接进行更新
		if i=layer and parentid<>"" then
		conn.execute("update class set child=child+1 where classid="&parentid)
		end if
	next
	'更新该版面排序以及大于本需要和同在本分类下的版面排序序号
	conn.execute("update class set ordersID=ordersID+1 where rootid="&rootid&" and ordersID>"&ordersID)
	conn.execute("update class set ordersID="&ordersID&"+1 where classid="&Request.form("newclassid"))
else
	'当上级分类深度为0的时候只要更新上级分类版面数和该版面排序序号即可
	conn.execute("update class set child=child+1 where classid="&request("parentID"))
	set rs=conn.execute("select max(ordersID) from class where classid="&Request.form("newclassid"))
	conn.execute("update class set ordersID="&rs(0)&"+1 where classid="&Request.form("newclassid"))
end if
end if
response.write "<p>类别添加成功！<br><br>"&str
call classinfo()
end sub

sub savedit()
if request("class")="" then
	Errmsg=Errmsg+"<br>"+"<li>请输入类别名称。"
	Founderr=true
end if
if request("parentID")="" then
	Errmsg=Errmsg+"<br>"+"<li>请选择父类或选择“没有父类”。"
	Founderr=true
end if
if founderr=true then
	response.write Errmsg
	exit sub
end if

if clng(request("editid"))=clng(request("parentID")) then
	response.write "所属类别不能指定自己"
	exit sub
end if
dim trs,brs,mrs
set rs = server.CreateObject ("adodb.recordset")
sql = "select * from class where classid="&request("editid")
rs.Open sql,conn,1,3
newclassid=rs("classid")
parentid=rs("parentid")
iparentid=rs("parentid")
ParentStr=rs("ParentStr")
layer=rs("layer")
child=rs("child")
rootid=rs("rootid")
'判断所指定的论坛是否其下属论坛
if ParentID=0 then
	if clng(request("parentID"))<>0 then
	set trs=conn.execute("select rootid from class where classid="&request("parentID"))
	if rootid=trs(0) then
		response.write "您不能指定该版面的下属论坛作为所属论坛"
		exit sub
	end if
	end if
else
	set trs=conn.execute("select classid from class where ParentStr like '%"&ParentStr&"%' and classid="&request("parentID"))
	if not (trs.eof and trs.bof) then
		response.write "您不能指定该版面的下属论坛作为所属论坛"
		response.end
	end if
end if
if parentid=0 then
	parentid=rs("classid")
	iparentid=0
end if
rs("class")=Request.Form("class")
rs("classtype") = Request.form("classtype")
rs.Update 
rs.Close
set rs=nothing

set mrs=conn.execute("select max(rootid) from class")
Maxrootid=mrs(0)+1
if clng(parentid)<>clng(request("parentID")) and not (iparentid=0 and cint(request("parentID"))=0) then
	'如果原来不是一级分类改成一级分类
	if iparentid>0 and cint(request("parentID"))=0 then
		'更新当前版面数据
		conn.execute("update class set layer=0,ordersID=0,rootid="&maxrootid&",parentid=0,parentstr='0' where classid="&newclassid)
		ParentStr=ParentStr & ","
		set rs=conn.execute("select count(*) from class where ParentStr like '%"&ParentStr&"%'")
		classcount=rs(0)
		if isnull(classcount) then
		classcount=1
		else
		classcount=classcount+1
		end if
		conn.execute("update class set child=child-"&classcount&" where classid="&iparentid)
		for i=1 to layer
			'得到其父类的父类的版面ID
			set rs=conn.execute("select parentid from class where classid="&iparentid)
			if not (rs.eof and rs.bof) then
				iparentid=rs(0)
				conn.execute("update class set child=child-"&classcount&" where classid="&iparentid)
			end if
		next
		if child>0 then
		i=0
		set rs=conn.execute("select * from class where ParentStr like '%"&ParentStr&"%'")
		do while not rs.eof
		i=i+1
		mParentStr=replace(rs("ParentStr"),ParentStr,"")
		conn.execute("update class set layer=layer-"&layer&",rootid="&maxrootid&",ParentStr='"&mParentStr&"' where classid="&rs("classid"))
		rs.movenext
		loop
		end if
	elseif iparentid>0 and cint(request("parentID"))>0 then
	set trs=conn.execute("select * from class where classid="&request("parentID"))
	ParentStr=ParentStr & ","
	set rs=conn.execute("select count(*) from class where ParentStr like '%"&ParentStr&"%'")
	classcount=rs(0)
	if isnull(classcount) then classcount=1
	conn.execute("update class set ordersID=ordersID + "&classCount&" + 1 where rootid="&trs("rootid")&" and ordersID>"&trs("ordersID")&"")
	conn.execute("update class set layer="&trs("layer")&"+1,ordersID="&trs("ordersID")&"+1,rootid="&trs("rootid")&",ParentID="&request("parentID")&",ParentStr='" & trs("parentstr") & "," & trs("classid") & "' where classid="&newclassid)
	i=1
	set rs=conn.execute("select * from class where ParentStr like '%"&ParentStr&"%' order by ordersID")
	do while not rs.eof
	i=i+1
	iParentStr=trs("parentstr") & "," & trs("classid") & "," & replace(rs("parentstr"),ParentStr,"")
	conn.execute("update class set layer=layer+"&trs("layer")&"-"&layer&"+1,ordersID="&trs("ordersID")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&iParentStr&"' where classid="&rs("classid"))
	rs.movenext
	loop
	ParentID=request("parentID")
	if rootid=trs("rootid") then
	conn.execute("update class set child=child+"&i&" where (not ParentID=0) and classid="&parentid)
	for k=1 to trs("layer")
		set rs=conn.execute("select parentid from class where (not ParentID=0) and classid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			conn.execute("update class set child=child+"&i&" where (not ParentID=0) and  classid="&parentid)
		end if
	next
	conn.execute("update class set child=child-"&i&" where (not ParentID=0) and classid="&iparentid)
	for k=1 to layer
		set rs=conn.execute("select parentid from class where (not ParentID=0) and classid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			conn.execute("update class set child=child-"&i&" where (not ParentID=0) and  classid="&iparentid)
		end if
	next
	else
	conn.execute("update class set child=child+"&i&" where classid="&parentid)
	for k=1 to trs("layer")
		set rs=conn.execute("select parentid from class where classid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			conn.execute("update class set child=child+"&i&" where classid="&parentid)
		end if
	next
	conn.execute("update class set child=child-"&i&" where classid="&iparentid)
	for k=1 to layer
		set rs=conn.execute("select parentid from class where classid="&iparentid)
		if not (rs.eof and rs.bof) then
			iparentid=rs(0)
			conn.execute("update class set child=child-"&i&" where classid="&iparentid)
		end if
	next
	end if 'end if rootid=trs("rootid") then
	else
	set trs=conn.execute("select * from class where classid="&request("parentID"))
	set rs=conn.execute("select count(*) from class where rootid="&rootid)
	classcount=rs(0)
	ParentID=request("parentID")
	conn.execute("update class set child=child+"&classcount&" where classid="&parentid)
	for k=1 to trs("layer")
		set rs=conn.execute("select parentid from class where classid="&parentid)
		if not (rs.eof and rs.bof) then
			parentid=rs(0)
			conn.execute("update class set child=child+"&classcount&" where classid="&parentid)
		end if
	next
	conn.execute("update class set ordersID=ordersID + "&classCount&" + 1 where rootid="&trs("rootid")&" and ordersID>"&trs("ordersID")&"")
	i=0
	set rs=conn.execute("select * from class where rootid="&rootid&" order by ordersID")
	do while not rs.eof
	i=i+1
	if rs("parentid")=0 then
		if trs("ParentStr")="0" then
		parentstr=trs("classid")
		else
		parentstr=trs("parentstr") & "," & trs("classid")
		end if
	conn.execute("update class set layer=layer+"&trs("layer")&"+1,ordersID="&trs("ordersID")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"',parentid="&request("parentID")&" where classid="&rs("classid"))
	else
		if trs("ParentStr")="0" then
		parentstr=trs("classid") & "," & rs("parentstr")
		else
		parentstr=trs("parentstr") & "," & trs("classid") & "," & rs("parentstr")
		end if
	conn.execute("update class set layer=layer+"&trs("layer")&"+1,ordersID="&trs("ordersID")&"+"&i&",rootid="&trs("rootid")&",ParentStr='"&ParentStr&"' where classid="&rs("classid"))
	end if
	rs.movenext
	loop
	end if
end if
response.write "<p>类别修改成功！<br><br>"&str
set rs=nothing
set mrs=nothing
set trs=nothing
call classinfo()
end sub

sub updateclassorders()
if not isnumeric(request("editID")) then
	response.write "非法的参数！"&request("editID")
	exit sub
end if
if request("uporders")<>"" and not Cint(request("uporders"))=0 then
	if not isnumeric(request("uporders")) then
	response.write "非法的参数！"
	exit sub
	elseif Cint(request("uporders"))=0 then
	response.write "请选择要提升的数字！"
	exit sub
	end if
	
	set rs=conn.execute("select classid,ParentID,ordersID,ParentStr,child from class where classid="&request("editID"))
	ParentID=rs(1)
	ordersID=rs(2)
	ParentStr=rs(3) & "," & request("editID")
	child=rs(3)
	i=0
	if child>0 then
	set rs=conn.execute("select count(*) from class where ParentStr like '%"&ParentStr&"%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	set rs=conn.execute("select classid,ordersID,child,ParentStr from class where ParentID="&ParentID&" and ordersID<"&ordersID&" order by ordersID desc")
	do while not rs.eof
	i=i+1
	if Cint(request("uporders"))>=i then
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select classid,ordersID from class where ParentStr like '%"&rs(3)&","&rs(0)&"%' order by ordersID")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		conn.execute("update class set ordersID="&ordersID&"+"&oldorders&"+"&ii&" where classid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		conn.execute("update class set ordersID="&ordersID&"+"&oldorders&" where classid="&rs(0))
		if Cint(request("uporders"))=i then 
		uporders=rs(1)
		end if
	end if
	ordersID=rs(1)
	rs.movenext
	loop
	'response.write "update class set ordersID="&uporders&" where classid="&request("editID")
	'response.end
	conn.execute("update class set ordersID="&uporders&" where classid="&request("editID"))
	if child>0 then
	i=uporders
	set rs=conn.execute("select classid from class where ParentStr like '%"&ParentStr&"%' order by ordersID")
	do while not rs.eof
	i=i+1
	conn.execute("update class set ordersID="&i&" where classid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
elseif request("doorders")<>"" then
	if not isnumeric(request("doorders")) then
	response.write "非法的参数！"
	exit sub
	elseif Cint(request("doorders"))=0 then
	response.write "请选择要下降的数字！"
	exit sub
	end if
	set rs=conn.execute("select ParentID,ordersID,ParentStr,child from class where classid="&request("editID"))
	ParentID=rs(0)
	ordersID=rs(1)
	ParentStr=rs(2) & "," & request("editID")
	child=rs(3)
	i=0
	if child>0 then
	set rs=conn.execute("select count(*) from class where ParentStr like '%"&ParentStr&"%'")
	oldorders=rs(0)
	else
	oldorders=0
	end if
	set rs=conn.execute("select classid,ordersID,child,ParentStr from class where ParentID="&ParentID&" and ordersID>"&ordersID&" order by ordersID")
	do while not rs.eof
	i=i+1
	if Cint(request("doorders"))>=i then
		if rs(2)>0 then
		ii=0
		set trs=conn.execute("select classid,ordersID from class where ParentStr like '%"&rs(3)&","&rs(0)&"%' order by ordersID")
		if not (trs.eof and trs.bof) then
		do while not trs.eof
		ii=ii+1
		'response.write "update class set ordersID="&ordersID&"+"&ii&" where classid="&trs(0)&"－a<br>"
		conn.execute("update class set ordersID="&ordersID&"+"&ii&" where classid="&trs(0))
		trs.movenext
		loop
		end if
		end if
		'response.write "update class set ordersID="&ordersID&" where classid="&rs(0)&"<br>"
		conn.execute("update class set ordersID="&ordersID&" where classid="&rs(0))
		if Cint(request("doorders"))=i then doorders=rs(1)
	end if
	ordersID=rs(1)
	rs.movenext
	loop
	'response.write "update class set ordersID="&doorders&" where classid="&request("editID")&"<br>"
	conn.execute("update class set ordersID="&doorders&" where classid="&request("editID"))
	'如果有下属论坛，则更新其下属论坛排序
	if child>0 then
	i=doorders
	set rs=conn.execute("select classid from class where ParentStr like '%"&ParentStr&"%' order by ordersID")
	do while not rs.eof
	i=i+1
	'response.write "update class set ordersID="&i&" where classid="&rs(0)&"－b<br>"
	conn.execute("update class set ordersID="&i&" where classid="&rs(0))
	rs.movenext
	loop
	end if
	'response.end
	set rs=nothing
	set trs=nothing
end if
'response.write "?action=orders&editid="&request("parentID")
response.redirect "?action=orders&editid="&request("parentID")
end sub
sub delclass()
conn.execute("delete from article where classid="&request("classid"))
response.write "清空该类数据成功，请返回更新数据！"
call classinfo()
end sub

sub del()
set rs=conn.execute("select ParentStr,child,layer from class where classid="&Request("editid"))
if not (rs.eof and rs.bof) then
if rs(1)>0 then
	response.write "该类别含有子类别，请删除其子类别后再进行删除本类别的操作"
	exit sub
end if
'如果有上级版面，则更新数据
if rs(2)>0 then
	conn.execute("update class set child=child-1 where classid in ("&rs(0)&")")
end if
sql = "delete from class where classid="&Request("editid")
conn.execute(sql)
conn.execute("delete from article where classid="&request("editid"))
end if
response.write "<P>类别删除成功！</P>"
call classinfo()
end sub
%>
  </center>
</div>
</body>
</html>
<%   set rs=nothing
     conn.close
     set conn=nothing
%>