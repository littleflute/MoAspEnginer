<%@ CODEPAGE = "936" %>


<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if
%>
<!--#include file="conn.asp"-->
<%
   username=LCase(Request("username"))
   password=LCase(Request("newpin"))
   right_class=CInt(Request("right_class"))

   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   rs.Open "Select * from Admin where username='"&username&"'",conn

   if not rs.EOF then

   Response.Write "<font color=red><div align=center><br><br>���û����Ѿ�����</div></font>"
   Response.End   

   end if
   rs.close
   sql="select * from admin"
   rs.open sql,conn,1,3
   rs.addnew
   rs("username")=username
   rs("password")=password
   rs("flag")=right_class
   rs.update
   rs.Close
   set rs=Nothing
   
   conn.Close
   set conn=Nothing

   Response.Redirect "adminuser.asp"
%>