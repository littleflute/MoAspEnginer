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
<P align="center">�ꡡ�⣺<INPUT size=85 name=title value=<%=title1%>></P>
<P align="center">�ڡ��ݣ�<TEXTAREA cols=73 name=body rows=15><%=content%></TEXTAREA></P>
<P align="center">�ؼ��֣�<INPUT size=25 name=keyw value=<%=keyw%>>
    ������ߣ�
<INPUT size=25 name=writer value=<%=write%>>�����ͣ�<select span style="font-size:10.5pt" name="type">
<option value="0">�ı�</option>
<option value="1">����</option>
</select>
<p align="center">
<INPUT class=buttonface type=submit value=" ȷ �� ">
<INPUT class=buttonface type=reset value=" �� �� "></p>
</form>
<P align="right"><a href="newsedit.asp">�༭����</a>��<a href=".">�鿴����</a>
</BODY>
</HTML>