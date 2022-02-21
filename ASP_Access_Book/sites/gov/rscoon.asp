<%
response.buffer=true	
dim conn
dim admin
dim connstr
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" &server.mappath("/gov.mdb")&";"

%>