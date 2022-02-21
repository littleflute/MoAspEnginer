<% Response.Buffer=True %>
<!--#include file="../inc/admin.inc"-->
<% set rs=server.createobject("adodb.recordset")
sql="update research set selecta=0,selectb=0,selectc=0,selectd=0 where id=1"
rs.open sql,conn,3,3
set rs=nothing
conn.close
set conn=nothing
response.write"<SCRIPT language=JavaScript>alert('调查引擎数据成功重置为0!');"
response.write"javascript:history.go(-1);</SCRIPT>" %>