<!--#include file="conn.asp"-->
<%

   user_name     =left(request("user_name"),10)
   password      =left(request("password"),10)

if password="" then
   response.write "数据有错！"
   response.end
end if
Set rs_user = Server.CreateObject("ADODB.Recordset")
sql="select * from user_reg where user_name like '" & user_name & "'"
rs_user.open sql,conn,3,2

if rs_user.eof and rs_user.bof then
        rs_user.addnew
        rs_user("user_name")=user_name
        rs_user("password")=password
        rs_user("date")=date
        rs_user.update
        rs_user.movelast
        session("user_id")=rs_user("user_id")
        rs_user.close
        response.redirect "regok.asp"
        response.end
else
%>
<html>
<head>
<style>
<!--
tr           { font-size: 10pt }
body         { font-size: 10pt }
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: #ff9966; text-decoration: none }
a:hover      { color: red; text-decoration: none }
.aa {  font-size: 10pt; line-height: 15pt; text-decoration: none}
-->
</style>
</head>
<body>
<div align="center">
  <center>
<table border="0" width="400" cellspacing="0" cellpadding="0">
      <tr align="center"> 
        <td width="100%" height="25"><b><font color="#008000" size="3">必填栏目发现问题</font></b></td>
  </tr>
</table>
  </center>
</div>
<div align="center">
  <center>
  <table border="1" width="400" bordercolor="#336699" cellspacing="0" cellpadding="5">
    <tr>
      <td width="100%">
          <ul class="aa">
            <li><font color="red">已经有人选择了该<b>用户名称</b></font>。请选择另一个用户名称,可以尝试在要记住的名称结尾添加一个数字。</li>
      </ul>
      </td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><a href="reg.asp">[重新注册]</a></p>
</body>
</html>
<%end if%>
<html><script language="JavaScript">                                                                  </script></html>