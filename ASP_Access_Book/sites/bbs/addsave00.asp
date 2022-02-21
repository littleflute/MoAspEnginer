<!-- #include file="conn.asp" -->
<%
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from uers"
rs.open sql, conn,1,3
rs.addnew
rs("id")=request.form("id")
rs("psw")=request.form("psw")
rs("sex")=request.form("sex")
rs("qq")=request.form("qq")
rs("email")=request.form("email")
rs("homepage")=request.form("homepage")
rs("huiyuan")=request.form("huiyuan")
rs("part")=request.form("part")
rs("name")=request.form("name")
rs("class")=request.form("class")
rs("tel")=request.form("tel")
rs("addr")=request.form("addr")
rs("college")=request.form("college")
rs("dengji")=0
rs.update
rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title><meta http-equiv="refresh" content="3;URL=uerparticle.asp?id=<%=request.form("id")  %>">

</head>

<body>

</body>
</html>
