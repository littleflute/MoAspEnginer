<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("data/picture.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
%>