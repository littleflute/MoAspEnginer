<%
dim conn,DBPath
Set conn = Server.CreateObject("ADODB.Connection")
conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" &server.mappath("data/data.mdb")&";"

%>