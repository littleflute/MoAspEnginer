<!-- #include file="../conn.asp" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>新闻</title>
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
	response.write "<p>还 没 有 任 何 新 闻</p>"
else
	if not isempty(request("page")) then
		pagecount=cint(request("page"))
	else
		pagecount=1
	end if
	rs.PageSize=10
	rs.AbsolutePage=pagecount
%>
    <p><strong>全部新闻</strong><table width=100%>
<%For iPage = 1 To rs.PageSize
If rs.EOF Then Exit For
id=rs("ID")%>
<tr><td><a href="javascript:NewsWindow(<%=id%>)"><u><%=rs("title")%></u></a>--<%=rs("write")%>【<%=rs("times")%>】
</td><td align=right><img src=note.gif><a href=newsadd.asp?id=<%=id%>>编辑</a>　<img src=del.gif><a href=newsedit.asp?del=<%=id%>>删除</a></td></tr>
<%
rs.movenext
Next
Response.Write "</table><p>共"&rs.recordcount&"条记录<br>"
if rs.PageCount>1 Then
 If pagecount<>1 Then
   Response.Write "〖<A HREF=newsedit.asp?Page=1&boardid="&request("boardid")&">首页</A>〗"
   Response.Write "〖<A HREF=newsedit.asp?Page="&(pagecount-1)&">前页</A>〗"
 End If
 If pagecount<>rs.PageCount Then
   Response.Write "〖<A HREF=newsedit.asp?Page="&(pagecount+1)&">后页</A>〗"
   Response.Write "〖<A HREF=newsedit.asp?Page="&rs.PageCount&">尾页</A>〗"
 End If
End If
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<p align=right><a href=newsadd.asp>添加新闻</a>　<a href=".">查看新闻</a>