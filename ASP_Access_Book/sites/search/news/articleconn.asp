<%
   dim conn   
   dim connstr
   on error resume next
     set conn=server.createobject("ADODB.CONNECTION")
conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & DefaultDir &server.mappath("downloadlu.mdb")&";"
%>

