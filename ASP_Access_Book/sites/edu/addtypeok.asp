<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
addtype = trim(request("addtype"))
if addtype = "" then
 response.write "<script>alert('请输入要添加的栏目名');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if
if len(addtype) > 5 then
 response.write "<script>alert('栏目名不得超过5个汉字');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

sql = "select * from type where type='"&addtype&"'"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if not (rs.bof and rs.eof) then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('数据库中已经有一个名为"&addtype&"的栏目了');history.go(-1);</script>"
 response.end
else
 rs.addnew
 rs("type")=addtype
 rs.update
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
response.write "<script>alert('添加成功');window.location.href='addtype.asp';</script>"
%>