<!-- #include file="conn.asp" -->
<%

from=request.form("from")
content=request.form("content")
title=request.form("title")
ttto= request("to")
'**********检查是否填写了所有项，如果不是侧自动返回申请页面
if  ttto=""  or  title ="" or  content=""  then
errmsg=errmsg & "请填写一定要打**的内容！\n或者\n"
end if
'**********检查用户名.如果有重复用户名侧自动返回申请页面
dim rsc,errmsg
Set rsc = Conn.Execute("select * from uers where id= '" & ttto & "'")
if    rsc.eof  then  errmsg=errmsg & "没有找到该用户！\n"

if errmsg<>"" then
    Conn.Close
    Set conn = nothing
    Set rsc = nothing
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if
'**********检查结束**********
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from duanx"
rs.open sql, conn,1,3
rs.addnew
rs("to")=ttto
rs("from")=from
rs("title")=title
rs("content")=content
rs("flag")="否"
rs.update
rs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>成功！！！</title><meta http-equiv="refresh" content="1;URL=duanadd.asp">

</head>

<body>
<div align="center">
  <p><font color="#FF0000" size="3"><strong>成功注册！</strong></font></p>
  <p>系统将在3秒内自动返回你加入的资料！</p>
  <p></p>
  <p><a href="duanadd.asp">如果你不想继续等待请点击这里<br>
    (或者你的浏览器无法自动返回) <br>
    </a> </p>
</div>
</body>
</html>
