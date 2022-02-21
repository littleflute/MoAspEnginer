<%@ CODEPAGE = "936" %>
<!--#include file=conn.asp-->
<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<html>
<%
   Set rs=Server.CreateObject("Adodb.RecordSet")
   
   sql="select * from admin  order by id"
   
   rs.Open sql,conn,1,1
%>
<head>
<link rel="stylesheet" href="style.css">
</head>
<BODY bgcolor="#39867B">
<div align="center"></div>
<table border="0" cellspacing="1" width="758" align=center bgcolor="#C6EBDE">
  <tr>
    <td width="100%"><img src="../images/top.jpg" width="758" height="114"></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF"><br>
        <br>
        <table border="0" cellspacing="1" width="758" align=center bgcolor="#C6EBDE">
          <tr>
            <td width="100%">账号：<%=rs("id")%></td>
          </tr>
          <tr>
            <td width="100%" bgcolor="#FFFFFF">等级：
                <%if rs("flag")=1 then%>
            超级用户
            <%end if%>
            <%if rs("flag")=2 then%>
            普通用户
            <%end if%>
            <%if rs("flag")=3 then%>
            员工
            <%end if%>
            </td>
            <b>管理员可进行操作说明</b>：<br>
            <br>
          1，对已经添加软件修改或删除，请点左边相关连接进行操作。操作用户：超级用户，普通管理员 <br>
          <br>
          2，对软件栏目进行添加修改删除，请点左边相关连接进行操作。操作用户：超级用户 <br>
          <br>
          3，用户的增加修改删除，请点左边相关连接进行操作。操作用户：超级用户 <br>
          <br>
          4，<font color=red>为了系统的安全性，离开管理请点击退出系统</font> <br>
          <br>
          </tr>
      </table></td>
  </tr>
</table>
</html>