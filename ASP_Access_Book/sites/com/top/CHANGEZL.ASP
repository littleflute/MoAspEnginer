
<!-- #include file="err.asp" -->
<% 
response.buffer=true
Response.Expires=0
%>
<!-- #include file="connect.asp" -->
<%
adminame=request("username")
adminpwd=request("password")
sql="select * from admin "
rs.open sql,conn,3,2
rs("name")=adminame
rs("password")=adminpwd
rs.update 
rs.close 
%>
<script language=vbscript>  
  		MsgBox "您成功修改了您的管理员帐号和密码，请牢记并重新登陆！"  
  		location.href = "login.htm"  
	  </script> 