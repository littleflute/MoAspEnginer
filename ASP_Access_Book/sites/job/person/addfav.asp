<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid")
fav=request("fav")
if fav<>"" then 
set rs=server.createobject("adodb.recordset")
sql="select * from company where uname='"&fav&"'"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
response.write"<SCRIPT language=JavaScript>alert('�û�["&fav&"]�����ڣ�');"
response.write"javascript:history.go(-1);</SCRIPT>"
end if
cname=rs("cname")
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from pfavorite where uname='"&uname&"'and fuid='"&fav&"'"
rs.open sql,conn,1,1
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('["&cname&"]����ӵ��ղؼУ������ظ���ӣ�');"
response.write"javascript:history.go(-1);</SCRIPT>"
end if
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from pfavorite"
rs.open sql,conn,3,3
rs.addnew
rs("uname")=uname
rs("fuid")=request("fav")
rs.update
rs.close
set rs=nothing
response.write"<SCRIPT language=JavaScript>alert('["&cname&"]�ɹ���ӵ��ղؼУ�');"
response.write"javascript:history.go(-1);</SCRIPT>"
end if %>