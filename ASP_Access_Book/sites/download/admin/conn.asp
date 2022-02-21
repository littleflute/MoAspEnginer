<%
dim conn
dim connstr
Set conn = Server.CreateObject("ADODB.Connection")
connstr="DBQ="+server.mappath("../download.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"

	conn.Open connstr
%>