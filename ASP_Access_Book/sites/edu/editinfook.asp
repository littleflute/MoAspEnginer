<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<!--#include file="md5.asp"-->
<!--#include file="isteacher.asp"-->
<%
fenlei1 = server.htmlencode(trim(request("fenlei1")))
fenlei2 = server.htmlencode(trim(request("fenlei2")))
teacher = server.htmlencode(trim(request("teacher")))
loginname = trim(request("loginname"))
email = server.htmlencode(trim(request("email")))
homepage = server.htmlencode(trim(request("homepage")))
address = server.htmlencode(trim(request("address")))
intro = server.htmlencode(trim(request("intro")))
photourl = server.htmlencode(trim(request("fileurl")))
qq = server.htmlencode(trim(request("qq")))
password = request("password")
password1 = request("password1")
ask = trim(request("ask"))
answer = trim(request("answer"))

if fenlei1 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������������"&strfenlei1&"');history.go(-1);</script>"
 response.end
end if
if fenlei2 = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������������"&strfenlei2&"');history.go(-1);</script>"
 response.end
end if
if teacher = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������������');history.go(-1);</script>"
 response.end
end if
if len(teacher) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������ó���10������');history.go(-1);</script>"
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
sql = "select count(teacherid) from teacher where loginname='"&loginname&"' and teacherid<>"&session("teacherid")
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

if len(email) > 30 then
 conn.close
 set conn = nothing
 response.write "<script>alert('E-Mail��ַ���ó���30���ַ�');history.go(-1);</script>"
 response.end
end if
if len(homepage) > 50 then
 conn.close
 set conn = nothing
 response.write "<script>alert('������ҳ���ó���50���ַ�');history.go(-1);</script>"
 response.end
end if
if len(qq) > 10 then
 conn.close
 set conn = nothing
 response.write "<script>alert('QQ���벻Ӧ�ó���10λ�ɣ�');history.go(-1);</script>"
 response.end
end if
if len(address) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('ͨѶ��ַ���ó���25������');history.go(-1);</script>"
 response.end
end if
if photourl <> "" and left(photourl,6) <> "files" and right(photourl,3) <> "jpg" and right(photourl,3) <> "gif" and right(photourl,3) <> "png" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��Ƭ��ʽ���ԣ�ֻ֧��jpg��gif��png��ʽ��ͼƬ');history.go(-1);</script>"
 response.end
end if
if intro = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��������˼��');history.go(-1);</script>"
 response.end
end if
if password <> password1 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����������벻һ��');history.go(-1);</script>"
 response.end
end if
if ask = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����������һ�����');history.go(-1);</script>"
 response.end
end if
if len(ask) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����һ����ⲻ�ó���25������');history.go(-1);</script>"
 response.end
end if
if answer = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����������һش�');history.go(-1);</script>"
 response.end
end if
if len(answer) > 25 then
 conn.close
 set conn = nothing
 response.write "<script>alert('�����һش𰸲��ó���25������');history.go(-1);</script>"
 response.end
end if

sql = "select * from teacher where teacherid="&session("teacherid")
rs.open sql,conn,1,3
if rs.bof and rs.eof then
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  response.write "<script>alert('�벻Ҫ����');top.window.location.href='teachermain.asp';</script>"
  response.end
else
  rs("teacher")=teacher
  rs("fenlei1")=fenlei1
  rs("fenlei2")=fenlei2
  rs("email")=email
  if homepage <> "" and left(homepage,7) <> "http://" then
   homepage = "http://"&homepage
  end if
  rs("homepage")=homepage
  rs("address")=address
  rs("photourl")=photourl
  rs("qq")=qq
  rs("intro")=intro
  rs("ask")=ask
  rs("answer")=answer
  rs("loginname")=loginname
  if password <> "" then rs("password")=password
  rs.update
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
end if
response.write "<script>alert('�޸ĳɹ�');window.location.href='editinfo.asp?id="&session("teacherid")&"';</script>"
%>