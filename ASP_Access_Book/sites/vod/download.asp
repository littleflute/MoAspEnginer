<!--#include file="conn.asp"-->
<%
set rs=server.createobject("adodb.recordset")
sql="select * from download where id="&request("id")         
rs.open sql,conn,1,1
%>
<%
if session("user")="" and rs("club")<>"" then
%>
<script>alert('�Բ��𣡴�Ϊ��ԱӰƬ����ע������ǵĻ�Ա��');window.close();</script>
<%else         
if request("downid")="2" then
url=""&rs("filename1")&""
elseif request("downid")="3" then
url=""&rs("filename2")&""
else
url=""&rs("filename")&""
end if

response.redirect ""&url&""
response.end
rs.close                               
set rs=nothing                               
conn.close                               
set conn=nothing%>
<%end if%>

















