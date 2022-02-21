<!--#include file="conn.asp"-->
<%
name = trim(request("name"))
if name = "" then
 response.write "<script>alert('请输入学生姓名');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if
if len(name) > 5 then
 response.write "<script>alert('学生姓名不得超过5个汉字');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

title = trim(request("title"))
if title = "" then
 response.write "<script>alert('请输入作业标题');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

message = trim(request("message"))
if message = "" then
 response.write "<script>alert('请输入作业答案');history.go(-1);</script>"
 conn.close
 set conn = nothing
 response.end
end if

reid = trim(request("reid"))
if reid = "" then
 response.write "<script>alert('非法操作');history.go(-1);</script>"
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
 response.write "<script>alert('你已经提交过作业了');history.go(-1);</script>"
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
response.write "<script>alert('添加成功');window.location.href='index.asp';</script>"
%>