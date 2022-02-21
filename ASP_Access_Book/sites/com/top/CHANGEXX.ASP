<!-- #include file="err.asp" -->
<% 
response.buffer=true
Response.Expires=0
%>
<!-- #include file="connect.asp" -->
<%
id=request.form("id")
name=request.form("name")
neirong=request.form("neirong")
rs.open "select * from data where id="&id,conn,3,2
rs("name")=name
rs("neirong")=neirong
rs.update 
rs.close 
%>
<script language=vbscript>  
  		MsgBox "您成功修改了标题和内容！"  
  		location.href = "add.asp"  
	  </script>  