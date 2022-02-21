<% Response.Buffer=True %>
<!--#include file="../inc/dbconn.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/index.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;人才市场—&gt;管理登录</title>
</head>
<form action=login.asp method=post>
<body topmargin="0" leftmargin="0">

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="390">
    <tr>
      <td width="413" height="10" valign="top"></td>
    </tr>
  </center>
  <tr>
    <td valign="top" height="380">
      <div align="center">
        <center>
        <table border="1" cellpadding="0" cellspacing="0" width="202" height="132" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
          <tr>
              <td height="17" width="198" background="../images/t-bg1.gif" valign="bottom"> 
                <p align="center">=== 管理员登录 ===</td>    
          </tr>
        </center>
          <tr>
              <td height="113" width="198" bgcolor="#fffff4">
<p align="center"><br>
              帐 号:<input type="text" name="admin" size="15" maxLength=15 style="font-family: 宋体; color: #000060; font-size: 9pt"><br> 
              <br>
              密 码:<input type="password" name="pwd" size="15" maxLength=15 style="font-family: 宋体; color: #000060; font-size: 9pt"> </p>            
        <center>
        <p><input type="submit" value="登 录" name="B1" style="font-family: 宋体; font-size: 9pt; height: 19; position: relative"><br>
                    <br>
        </center></td>
          </tr>
        </table>
      </div>
      <p>　
    </td>
  </tr>
  </table>
</div>

</body>

</html>
<% admin=request("admin")
   if admin="" then
   response.end
   end if
   pwd=request("pwd")
   set rs=server.createobject("adodb.recordset")
   sql="select * from admin where admin='"&admin&"' and pwd='"&pwd&"'"
   rs.open sql,conn,1,1
   if rs.bof or rs.eof then
   response.write"<SCRIPT language=JavaScript>alert('错误的用户或名密码，请重新输入！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   else
   session("flag")=admin
   response.Redirect "mnews.asp"
   end if %>

