<!--#include file="conn.asp"-->
<%
name = trim(request("name"))
if name = "" then
 response.write "<script>alert('������ѧ������');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if
if len(name) > 5 then
 response.write "<script>alert('ѧ���������ó���5������');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

title = trim(request("title"))
if title = "" then
 response.write "<script>alert('��������ҵ����');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

message = trim(request("message"))
if message = "" then
 response.write "<script>alert('��������ҵ��');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

reid = trim(request("reid"))
if reid = "" then
 response.write "<script>alert('�Ƿ�����');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

sql = "select * from work where name='"&name&"' and reid='"&reid&"'"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not (rs.bof and rs.eof) then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('���Ѿ��ύ����ҵ��');history.go(-1);</script>"
 response.end
else
 rs.addnew
 rs("reid")=reid
 rs("name")=name
  rs("title")=title
   rs("message")=message
 rs.update
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
response.write "<script>alert('��ӳɹ�');window.location.href='index.asp';</script>"
%>