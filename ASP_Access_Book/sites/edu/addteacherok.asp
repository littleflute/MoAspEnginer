<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<%
fenlei1 = server.htmlencode(trim(request("fenlei1")))
fenlei2 = server.htmlencode(trim(request("fenlei2")))
teacher = server.htmlencode(trim(request("teacher")))
loginname = trim(request("loginname"))
email = server.htmlencode(trim(request("email")))
homepage = server.htmlencode(trim(request("homepage")))
address = server.htmlencode(trim(request("address")))
qq = server.htmlencode(trim(request("qq")))
intro = server.htmlencode(trim(request("intro")))
password = request("password")
password1 = request("password1")
ask = trim(request("ask"))
answer = trim(request("answer"))

if fenlei1 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������ʦ���ڵ�"&strfenlei1&"');history.go(-1);</script>"
 response.end
end if
if fenlei2 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������ʦ���ڵ�"&strfenlei2&"');history.go(-1);</script>"
 response.end
end if
if teacher = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������ʦ����');history.go(-1);</script>"
 response.end
end if
if len(teacher) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('��ʦ�������ó���10������');history.go(-1);</script>"
 response.end
end if
if loginname = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������¼��');history.go(-1);</script>"
 response.end
end if
if len(loginname) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('��¼�����ó���10������');history.go(-1);</script>"
 response.end
end if
if IsNumeric(loginname) then
 conn.close
 set conn = nothing
 response.write "<script>alert('��¼������ȫΪ����');history.go(-1);</script>"
 response.end
end if

set rs = server.createobject("adodb.recordset")
sql = "select count(teacherid) from teacher where loginname='"&loginname&"'"
rs.open sql,conn,1,1
if rs(0) > 0 then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('�õ�¼���Ѿ�����ʹ����');history.go(-1);</script>"
 response.end
end if
rs.close

sql = "select * from teacher where teacherid is null"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,3
if  (rs.bof and rs.eof) then
 rs.addnew
 rs("teacher")=teacher
 rs("password")=password
 rs("fenlei1")=fenlei1
 rs("fenlei2")=fenlei2
 rs("email")=email
 rs("address")=address
 rs("intro")=intro
 rs("homepage")=homepage
 rs("qq")=qq
 rs("answer")=answer
 rs("ask")=ask
 rs("loginname")=loginname
 rs.update
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
end if
response.write "<script>alert('��ӳɹ�');window.location.href='adminsearchteacher.asp';</script>"
%>