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
 response.write "<script>alert('�������ʦ����');history.go(-1);</script>"
 response.end
end if
if course = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������������');history.go(-1);</script>"
 response.end
end if
if len(course) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������Ʋ��ó���25������');history.go(-1);</script>"
 response.end
end if
if fileurl = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������ϵ�ַ');history.go(-1);</script>"
 response.end
end if
if len(fileurl) > 100 then
 conn.close
 set conn = nothing
 response.write "<script>alert('���ϵ�ַ���ó���100��Ӣ����ĸ');history.go(-1);</script>"
 response.end
end if
if title = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������ϱ���');history.go(-1);</script>"
 response.end
end if
if len(title) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('���ϱ��ⲻ�ó���25������');history.go(-1);</script>"
 response.end
end if
if typeid = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��ѡ����������');history.go(-1);</script>"
 response.end
end if
if filesize < 1 then filesize = 0
if content = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������');history.go(-1);</script>"
 response.end
end if

sql = "select * from teacher where teacher='"&teacher&"'"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.eof then
 response.write "<script>alert('����Ľ�ʦ������');history.go(-1);</script>"
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
response.write "<script>alert('�����ɹ�');window.location.href='pub.asp';</script>"
%>