<!--#include file="conn.asp"-->

<%
dim id
id=trim(request("id"))

set rs=server.createobject("adodb.recordset")
sql="select * from pic where id="& id
rs.open sql,conn,1,3

rs("pic_count")=rs("pic_count")+1
rs.update
%>

<p><img border="0" src="<%=rs("pic_url")%>"></p>
<%
rs.close
set rs=nothing
%>
