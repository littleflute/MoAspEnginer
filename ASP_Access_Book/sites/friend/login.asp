<%
Option Explicit
dim rs_user
dim sql
dim user_name,password

   user_name     =left(request("user_name"),10)
   password      =left(request("password"),10)

if password="" or user_name="" then
   response.write "<script language='javascript'>"
   response.write "alert('数据填写有错!');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
end if

if InStr(LCase(password),"'")<>0 or InStr(LCase(password),"or")<>0 then 
		response.write "<script language='javascript'>"
		response.write "alert('密码不合法，请重新输入!');"
		response.write "history.go(-1);"
		response.write "</script>"
		response.end
		end if

if server.HTMLEncode(user_name)<>user_name or InStr(user_name,"【")<>0 or InStr(user_name,"】")<>0 or InStr(user_name," ")<>0 or InStr(user_name,"　")<>0 or InStr(user_name,"")<>0 then 
response.write "<script language='javascript'>"
   response.write "alert('数据填写有错!');"
   response.write "history.go(-1);"
   response.write "</script>"
   response.end
   end if
%>
<!--#include file="conn.asp"-->
<%
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql="select * from user_reg where user_name like '" & user_name & "'and password like '" & password & "'"
rs_user.open sql,conn,3,2

if rs_user.eof and rs_user.bof then
%>
<html>

<head>
<style>
<!--
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: #ff9966; text-decoration: none }
a:hover      { color: red; text-decoration: none }
-->
</style>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>提示信息</title>
</head>

<body>

<div align="center">
  <center>
    <table border="1" width="400" bordercolor="#CC0000" cellspacing="0" cellpadding="0" height="120" bgcolor="#B89D77">
      <tr>
        <td width="100%" height="16"> 
          <p align="center"><b><font color="#FFFFFF" size="3">错误提示</font></b>
        </td>
    </tr>
      <tr> 
        <td width="100%" bgcolor="#FFEBE1" height="100"> 
          <p align="center"><font size="3">用户名或密码错误！</font></p>
        <p align="center"><font size="2"><a href="default.asp">[返回]</a></font></td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>
<%else
        session("user_id")=rs_user("user_id")
        rs_user.close
        set rs_user=nothing
        set conn=nothing
        response.redirect "your.asp"
        response.end
end if%>        