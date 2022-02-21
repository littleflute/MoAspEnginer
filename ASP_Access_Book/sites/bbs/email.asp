<!-- #include file="conn.asp" -->

<%
 set rs= conn.execute("select email from uers")

 while not rs.eof
%>
<%=rs("email") %><br>
<%
 wend
%>