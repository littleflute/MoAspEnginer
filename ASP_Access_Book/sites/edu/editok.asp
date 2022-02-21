<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<!--#include file="fenlei.asp"-->
<%
'on error resume next
course = server.htmlencode(trim(request("course")))
fileurl = server.htmlencode(trim(request("fileurl")))
content = server.htmlencode(trim(request("content")))
title = server.htmlencode(trim(request("title")))
typeid = trim(request("type"))
filesize = int(trim(request("filesize")))

id = request("id")

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');top.window.location.href='teachermain.asp';</script>"
 response.end
end if
if course = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入资料名称');history.go(-1);</script>"
 response.end
end if
if len(course) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('资料名称不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if fileurl = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入资料地址');history.go(-1);</script>"
 response.end
end if
if len(fileurl) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('资料地址不得超过100个英文字母');history.go(-1);</script>"
 response.end
end if
if title = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入资料标题');history.go(-1);</script>"
 response.end
end if
if len(title) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('资料标题不得超过25个汉字');history.go(-1);</script>"
 response.end
end if
if typeid = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请选择资料类型');history.go(-1);</script>"
 response.end
end if
if filesize < 1 then filesize = 0
if content = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入简介');history.go(-1);</script>"
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
 response.write "<script>alert('请不要捣乱');top.window.location.href='teachermain.asp';</script>"
 response.end
else
 if rs("idofteacher") <> int(session("teacherid")) and session("admin") <> "admin" then
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  response.write "<script>alert('这个资料不是你发布的，你想干什么？');top.window.location.href='teachermain.asp';</script>"
  response.end
 else
  rs("fileurl")=fileurl
  rs("course")=course
  rs("dateandtime")=now()
  rs("content")=content
  rs("title")=title
  rs("idoftype")=cint(typeid)
  rs("filesize")=filesize
  rs.update
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
 end if
end if

response.write "<script>alert('修改成功');window.location.href='edit.asp?id="&id&"';</script>"
%>