<!-- #include file="../conn.asp" -->
<html>

<head>
<title>新闻</title>
<link rel="stylesheet" type="text/css" href="news.css">
</head>
<%
id=trim(request("id"))
set rs=server.createobject("adodb.recordset")
sql="SELECT * from news where ID="&id
rs.open sql,conn,1,1
if rs.eof then
	response.write "错误的ID号"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else%>
<table width="100%" HEIGHT=6% border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#FAD185"> 
    <td height=28 align=center bgcolor="#00345A"><b><font color="#ffffff"><%=rs("title")%></font></b></td>
  </tr>
  <tr> 
    <td height="344" valign=top> 
      <%if rs("type")=0 then%>
      <p><%=replace(rs("content"),chr(13)&chr(10),"<p>")%> 
        <%else%>
        <iframe id=BoardTitle name=newscon style="HEIGHT:100%; VISIBILITY: inherit; WIDTH:100%; Z-INDEX: 2" scrolling=auto frameborder=0 src="<%=rs("content")%>" ></iframe>
    </td>
<%end if%>
</tr>
</table>
<br>
<b>相关新闻：</b>
<ul>
<%
sql="SELECT * from news where ID<>"&id&" and keyw='"&rs("keyw")&"' order by ID desc"
rs.close
rs.open sql,conn,1,1
do while not rs.eof
%>
<li><a href=newswind.asp?id=<%=rs("ID")%>><u><%=rs("title")%></u></a>--<%=rs("write")%>【<%=rs("times")%>】
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%></ul>
</html>