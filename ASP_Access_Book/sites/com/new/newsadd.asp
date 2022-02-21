<!-- #include file="../conn.asp" -->
<link rel="stylesheet" type="text/css" href="news.css">
<%
id=Request("id")
if Request.form("title")<>"" then
title=Trim(request.form("title"))   
sql="select * from news where "
	if id<>"" then
		sql=sql&"id="&id
	else
		sql=sql&"title='"&title&"'"
	end if
set rs=Server.CreateObject("ADODB.recordset")
rs.Open sql,conn,1,3
if rs.eof or rs.bof then
	rs.addnew
end if
rs("title")=Request.form("title")
rs("content")=Request.form("body")
rs("type")=Request.form("type")
rs("keyw")=Request.form("keyw")
'rs("write")=Request.form("writer")
rs.Update

end if     

if id<>"" then
	sql="select * from news where id="&id
	set rs=Server.CreateObject("ADODB.recordset")
	rs.Open sql,conn,1,1
	if not rs.eof then
			title1=rs("title")
			content=rs("content")
			keyw=rs("keyw")
			write=rs("write")
	end if
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
end if
%>

<form name=form1 method="post" action="newsadd.asp">
<input value="<%=id%>" type=hidden name=id>
<P align="center">标　题：<INPUT size=85 name=title value=<%=title1%>></P>
<P align="center">内　容：<TEXTAREA cols=73 name=body rows=15><%=content%></TEXTAREA></P>
<P align="center">关键字：<INPUT size=25 name=keyw value=<%=keyw%>>
    　添加者：
<INPUT size=25 name=writer value=<%=write%>>　类型：<select span style="font-size:10.5pt" name="type">
<option value="0">文本</option>
<option value="1">链接</option>
</select>
<p align="center">
<INPUT class=buttonface type=submit value=" 确 定 ">
<INPUT class=buttonface type=reset value=" 清 除 "></p>
</form>
<P align="right"><a href="newsedit.asp">编辑新闻</a>　<a href=".">查看新闻</a>
</BODY>
</HTML>