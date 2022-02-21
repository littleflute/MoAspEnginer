<%@ Language=VBScript %>
<!--#include file=conn.asp -->
<%
AdminUID=trim(Request.Form("AdminUID"))
AdminPWD=trim(Request.Form("AdminPWD"))

set rs=server.CreateObject ("ADODB.RecordSet")
sql="select [id],psw,[flag],face,lasttime,loginnum from uers where [id]='" & AdminUID & "' and  psw='" & AdminPWD & "'"

rs.Open sql,conn,1,1
if not rs.eof then 
   session("AdminUID")=rs(0)
   session("AdminPWD")=rs("psw")
   session("face")=rs("face")
   if not rs("flag")="" then  
session("Adminflag")=rs("flag") 
else 
session("Adminflag")="" end if
rs("lasttime")=now()
rs("loginnum")=rs("loginnum")+1
rs.update
rs.close
set rs=nothing
'更新用户登陆次数和登陆时间


conn.close
set conn=nothing
    Response.Redirect "loginCheckok.ASP"
else
rs.close
set rs=nothing
set conn=nothing
set rs=nothing
   call usererror
end if





sub usererror 
    Response.Redirect "loginCheckerr.ASP"
end sub
%>


