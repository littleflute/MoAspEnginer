<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
id = trim(request("id"))
addtype = trim(request("addtype"))
if addtype = "" then
 response.write "<script>alert('��������Ŀ��');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if
if len(addtype) > 5 then
 response.write "<script>alert('��Ŀ�����ó���5������');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

sql = "select * from type where type='"&addtype&"' and typeid<>"&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if not (rs.bof and rs.eof) then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('���ݿ����Ѿ���һ����Ϊ"&addtype&"����Ŀ��');history.go(-1);</script>"
 response.end
end if
rs.close
set rs = nothing
conn.execute "update type set type='"&addtype&"' where typeid="&id
conn.close
set conn = nothing
response.write "<script>alert('�޸ĳɹ�');window.location.href='addtype.asp';</script>"
%>