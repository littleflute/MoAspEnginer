<!-- #include file="../conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����</title>
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
if request("del")<>"" then conn.Execute("delete  from news where id="&request("del"))
sql="SELECT * from news order by ID desc"
rs.open sql,conn,1,1
if rs.eof and rs.bof then
	response.write "<p>�� û �� �� �� �� ��</p>"
else
	if not isempty(request("page")) then
		pagecount=cint(request("page"))
	else
		pagecount=1
	end if
	rs.PageSize=10
	rs.AbsolutePage=pagecount
%>
    <p><strong>ȫ������</strong><table width=100%>
<%For iPage = 1 To rs.PageSize
If rs.EOF Then Exit For
id=rs("ID")%>
<tr><td><a href="javascript:NewsWindow(<%=id%>)"><u><%=rs("title")%></u></a>--<%=rs("write")%>��<%=rs("times")%>��
</td><td align=right><img src=note.gif><a href=newsadd.asp?id=<%=id%>>�༭</a>��<img src=del.gif><a href=newsedit.asp?del=<%=id%>>ɾ��</a></td></tr>
<%
rs.movenext
Next
Response.Write "</table><p>��"&rs.recordcount&"����¼<br>"
if rs.PageCount>1 Then
 If pagecount<>1 Then
   Response.Write "��<A HREF=newsedit.asp?Page=1&boardid="&request("boardid")&">��ҳ</A>���@"
   Response.Write "��<A HREF=newsedit.asp?Page="&(pagecount-1)&">ǰҳ</A>���@"
 End If
 If pagecount<>rs.PageCount Then
   Response.Write "��<A HREF=newsedit.asp?Page="&(pagecount+1)&">��ҳ</A>���@"
   Response.Write "��<A HREF=newsedit.asp?Page="&rs.PageCount&">βҳ</A>��"
 End If
End If
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<p align=right><a href=newsadd.asp>�������</a>��<a href=".">�鿴����</a>