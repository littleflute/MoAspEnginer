<%@LANGUAGE="VBSCRIPT"%>
<%
	'option explicit
	dim conn
	dim connstr
	dim db

	Set conn = Server.CreateObject("ADODB.Connection")
conn.Open  "PROVIDER=SQLOLEDB;DATA SOURCE=(local);UID=sa;PWD=password;DATABASE=download" 

function CloseDatabase

	Conn.close
	Set conn = Nothing

End Function
%>