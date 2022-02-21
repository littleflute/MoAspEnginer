<!--#include file="conn.asp"-->
<%
id = trim(request("id"))
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('没有找到您要阅读的资料');window.close();</script>"
 response.end
end if
sql = "select * from main where mainid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('没有找到您要阅读的资料');window.close();</script>"
 response.end
else
 rs("times") = rs("times") + 1
 rs.update
 response.redirect rs("fileurl")
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
%>