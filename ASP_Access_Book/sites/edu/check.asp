<!--#include file="conn.asp"-->
<!--#include file="md5.asp"-->
<%
id = trim(request("id"))
password = request("password")
verifycode = request("verifycode")

if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('���������ĵ�¼����ID');history.go(-1);</script>"
 response.end
end if
if password = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('����������');history.go(-1);</script>"
 response.end
end if

if session("verifycode") = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('��½��ʱ�������µ�¼');window.location.href='login.asp';</script>"
 response.end
elseif session("verifycode") <> verifycode then
 response.write "<script>alert('��֤���������');history.go(-1);</script>"
 response.end
end if

if IsNumeric(id) then
	sql = "select * from teacher where teacherid="&id
else
	sql = "select * from teacher where loginname='"&id&"'"
end if
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('��¼����ID�������û�иý�ʦ');history.go(-1);</script>"
else
 if password = rs("password") then
  if rs("locked") = 1 then
   rs.close
   set rs = nothing
   conn.close
   set conn = nothing
   response.write "<script>alert('���ѱ���ֹ��½���������Ա��ϵ');history.go(-1);</script>"
  else
   session("teacherid") = cstr(rs("teacherid"))
   rs.close
   set rs = nothing
   conn.close
   set conn = nothing
   response.redirect "teachermain.asp"
  end if
 else
  rs.close
  set rs = nothing
  conn.close
  set conn = nothing
  session("verifycode") = ""
  response.write "<script>alert('�������');window.location.href='login.asp';</script>"
 end if
end if
%>