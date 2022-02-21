<!--#include file="conn.asp"-->
<%
dim pic_name,pic_type,pic_class,pic_text,pic_url

set rs=server.createobject("adodb.recordset")
sql="select * from pic"
rs.open sql,conn,1,3
rs.addnew
rs("pic_name")=trim(request("pic_name"))
rs("pic_type")=trim(request("pic_type"))
rs("pic_class")=trim(request("pic_class"))
rs("pic_text")=replace(request("pic_text"),chr(10),"<br>")
rs("pic_url")=trim(request("pic_url"))
rs("pic_lurl")=trim(request("pic_lurl"))
rs.update
rs.close
set rs=nothing
response.redirect "add_pic.asp"
%>
