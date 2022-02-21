<!-- #include file="conn.asp" -->
<%
id=request.form("id")
psw=request.form("psw")
sex=request.form("sex")
qq=request.form("qq")
email=request.form("email")
homepage=request.form("homepage")
huiyuan=request.form("huiyuan")
part=request.form("part")
name=request.form("name")

tel=request.form("tel")
addr=request.form("addr")
college=request.form("college")
dengji=0

sql="INSERT INTO uers (id,psw,sex,qq,email,homepage,huiyuan,part,name,class,tel,addr,college,dengji) VALUES "
sql=sql & "(" & "'"& id & "'" & ","
sql=sql & "'"& psw &"'"&","
sql=sql & "'"& sex &"'"&","
sql=sql & "'"& qq &"'"&","
sql=sql & email &","
sql=sql & homepage & ","
sql=sql & huiyuan &","
sql=sql & part &","
sql=sql & name &","
sql=sql & tel &","
sql=sql & addr &","
sql=sql & college &","
sql=sql & "'" & dengji &"')"
conn.Execute (sql)
conn.close
set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title><meta http-equiv="refresh" content="3;URL=uerparticle.asp?id=<%=request.form("id")  %>">

</head>

<body>
加入成功！！！！！
</body>
</html>