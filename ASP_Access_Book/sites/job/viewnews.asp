<% Response.Buffer=True %> 
<!--#include file="inc/dbconn.inc"-->
<% set rs=server.createobject("adodb.recordset")
sql1="update jobnews set click=click+1 where id="&request("id")
rs.open sql1,conn,1,1
sql2="select * from jobnews where id="&request("id")
rs.open sql2,conn,1,1 %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>新闻资讯==><%=rs("title")%></title>
</head>

<body>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="369" height="35" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td width="367" height="17" valign="bottom" background="images/t-bg1.gif">
        <p align="center">=== <%=rs("title")%> ===</td>
    </tr>
    <tr>
        <td width="367" height="17" valign="top" bgcolor="#fffff4" align="center"> 
          <p align="left">&nbsp;&nbsp;&nbsp;<%=rs("text")%></td>
    </tr>
  </table>
  </center>
</div>

<p align="center">【<a href="javascript:window.close()">关闭窗口</a>】</p>

</body>

</html>
