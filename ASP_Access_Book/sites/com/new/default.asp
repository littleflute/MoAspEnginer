<!-- #include file="../conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>NEWS</title>
<link rel="stylesheet" type="text/css" href="news.css">
<script language="JavaScript">
function NewsWindow(id)
{
window.open('newswind.asp?id='+id,'infoWin', 'height=400,width=600,scrollbars=yes,resizable=yes');
}
</script>
</head>

<body>
<%
set rs=server.createobject("adodb.recordset")
sql="SELECT * from news order by ID desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<p>�� û �� �� �� �� ��</p>"
else%>
<p> 
  <%
cc=1
	if not isempty(request("page")) then
		pagecount=cint(request("page"))
	else
		pagecount=1
	end if
	rs.PageSize=10
	rs.AbsolutePage=pagecount
For iPage = 1 To rs.PageSize
If rs.EOF Then Exit For
if cc mod 2=1 then
	Response.Write "<tr bgcolor=#E7E7E7>"
else
	Response.Write "<tr BGCOLOR=#F4F4F4>"
end if
%>
  <img src=dot.gif><a href="javascript:NewsWindow(<%=rs("ID")%>)"><u><%=rs("title")%></u></a> <font size="1"><%=rs("times")%></font> 
  <br>
  <%
if DateDiff("d",rs("times"),date())<1 then Response.Write ""
Response.Write "</td></tr>"
cc=cc+1
rs.movenext
Next
Response.Write "</table><p>��"&rs.recordcount&"������&nbsp;"
if rs.PageCount>1 Then
 If pagecount<>1 Then
   Response.Write "<A HREF=default.asp?Page=1>��ҳ&nbsp;</A>" 
   Response.Write "<A HREF=default.asp?Page="&(pagecount-1)&">ǰҳ&nbsp;</A>"
 End If
 If pagecount<>rs.PageCount Then
   Response.Write "<A HREF=default.asp?Page="&(pagecount+1)&">��ҳ&nbsp;</A>"
   Response.Write "<A HREF=default.asp?Page="&rs.PageCount&">βҳ&nbsp;</A>"
 End If
End If
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</p>
<p><a href="newsedit.asp" target="_blank">�༭����</a></p>
</body>
</html>
