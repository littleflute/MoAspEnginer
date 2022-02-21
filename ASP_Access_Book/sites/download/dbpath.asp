<%
   'on error resume next
	dim conn
	dim dbpath
   	set conn=server.createobject("adodb.connection")
	DBPath = Server.MapPath("www.hooji.com.mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>