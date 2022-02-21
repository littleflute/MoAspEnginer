<%
sqlfenlei="select * from config"
set rsfenlei = server.createobject("adodb.recordset")
rsfenlei.open sqlfenlei,conn,1,1
strfenlei1 = rsfenlei("strfenlei1")
strfenlei2 = rsfenlei("strfenlei2")
rsfenlei.close
set rsfenlei = nothing
%>