<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<!--#include file="fenlei.asp"-->
<%
on error resume next
course = server.htmlencode(trim(request("course")))
fileurl = server.htmlencode(trim(request("fileurl")))
content = server.htmlencode(trim(request("content")))
title = server.htmlencode(trim(request("title")))
typeid = trim(request("type"))
filesize = int(trim(request("filesize")))
teacher = trim(request("teacher"))
if teacher = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请输入教师名称');history.go(-1);</script>"
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

sql = "select * from teacher where teacher='"&teacher&"'"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof then
 response.write "<script>alert('输入的教师不存在');history.go(-1);</script>"
 response.end
else
session("teacherid")=rs("teacherid")
end if
rs.close

sql = "select * from main"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
rs.addnew
rs("fileurl")=fileurl
rs("idofteacher")=int(session("teacherid"))
rs("course")=course
rs("dateandtime")=now()
rs("content")=content
rs("times")=0
rs("title")=title
rs("idoftype")=cint(typeid)
rs("filesize")=filesize
rs.update
rs.close
set rs = nothing
conn.close
set conn = nothing
response.write "<script>alert('发布成功');window.location.href='pub.asp';</script>"
%>