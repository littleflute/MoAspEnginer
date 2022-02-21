<%
set conn=Server.CreateObject("ADODB.connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("ilusw.mdb")
%>