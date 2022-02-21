<% Response.Buffer=True %>
<!--#include file="inc/dbconn.inc"--> 
<!--#include file="inc/enpasswd.inc"-->
<% uname=request("uname")
   pwd=mistake(request("pwd"))
   usertype=request("usertype")
   
   if usertype="person" then
   
   set rs=server.createobject("adodb.recordset")
   sql="select * from person where uname='"&uname&"' and pwd='"&pwd&"'"
   rs.open sql,conn,3,3
   if rs.bof or rs.eof then
   response.write"<SCRIPT language=JavaScript>alert('错误的用户名或密码，请重新输入！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   else
   session("puid")=uname
   response.Redirect "person/main.asp"
   end if 
   
   else
   
   set rs=server.createobject("adodb.recordset")
   sql="select * from company where uname='"&uname&"' and pwd='"&pwd&"'"
   rs.open sql,conn,3,3
   if rs.eof then
   response.write"<SCRIPT language=JavaScript>alert('错误的用户名或密码，请重新输入！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   else
   session("cuid")=uname 
   response.Redirect "company/main.asp"
   end if
   end if %>