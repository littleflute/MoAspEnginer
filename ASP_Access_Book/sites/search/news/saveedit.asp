<%
if request.cookies("adminok")="" then
  response.redirect "login.asp"
end if
%>
<!--#include file="articleconn.asp"-->
<!--#include file="inc/articlechar.inc"-->
<%

dim typename
dim title
dim content
dim sql
dim rs
dim filename
dim articleid
dim outfile
dim big
dim from
dim fromurl
dim url

title=htmlencode2(request.form("txttitle"))
url=htmlencode2(request.form("txturl"))
content=htmlencode2(request.form("txtcontent"))
from=htmlencode2(request.form("from"))
fromurl=htmlencode2(request.form("fromurl"))
big=htmlencode2(request.form("big"))
articleid=request("id")

set rs=server.createobject("adodb.recordset")
sql="select * from learning where articleid="&articleid
rs.open sql,conn,3,3
rs("title")=title
rs("url")=url
rs("content")=content
rs("big")=big
rs("from")=from
rs("fromurl")=fromurl
rs("dateandtime")=date()
rs.update
rs.close
set rs=noting
conn.close
set conn=nothing

response.redirect "manage.asp"
%>

