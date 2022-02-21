<!-- #include file="../conn.asp" -->
<%
set rs=server.createobject("adodb.recordset")
sql="SELECT top 18 * from news order by ID desc"
rs.open sql,conn,1,1
%>
document.write('<ul>')
<%
if rs.eof or rs.bof then%>
	document.write('还没有任何新闻')
<%else
while not rs.eof
OutPut="<li><a href=top/dismemo.asp?id="&rs("ID")&" target=_blank><u>"&rs("title")&"</u></a>"&rs("write")&"<font color=#666666>("&Month(rs("times"))&"-"&Day(rs("times"))&")</font></li>"
if DateDiff("d",rs("times"),date())<1 then OutPut=OutPut&"<font color=ff0000>new</font>"
%>
document.write('<%response.write OutPut%>')
<%
rs.movenext
wend
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
document.write('</ul>')