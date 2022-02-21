<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<!--#include file="../inc/html.inc"-->
<% uname=session("puid")
   reid=request("reid")
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"' and  job<>'""' "
rs.open sql,conn,1,1 
if rs.bof or rs.eof then
response.write"<SCRIPT language=JavaScript>alert('请先登录个人求职简历！');"
response.write"javascript:window.close();</SCRIPT>"
end if
sendname=rs("iname") 
rs.close                                                                    
set rs=nothing
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&reid&"'"
rs.open sql,conn,1,1 
if rs.bof or rs.eof then
response.write"<SCRIPT language=JavaScript>alert('用户["&reid&"]不存在！');"
response.write"javascript:window.close();</SCRIPT>"
end if
rename=rs("cname") 
rs.close
set rs=nothing %>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>发送站内信件</title>
</head>

<body>
<form method=post action="sendmail.asp?reid=<%=reid%>">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="408" height="231" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="15" width="406" valign="bottom" background="../images/t-bg1.gif">
        <p align="center">=== 发送站内信件 ===</td>     
    </tr>
	<tr>
        <td height="216" width="406" valign="top" bgcolor="#fffff4"> 
          <p align="center"><br>
        收信单位:<%=rename%></p>
        <p align="center">标 题:      
	  <input type="text" name="title" size="28" maxlength="20"></p>
        <p align="center">正 文:<br>     
        <textarea rows="14" name="mailtext" cols="44"></textarea>
          <p align="center"><input type="submit" value="确 定" style="position: relative; height: 19"><br>
      <br>
        </p>
      </td>
    </tr>
  </center>
  </table>
</div>

<p align="left">　</p>
</FORM>
</body>
</html>
<% mailtext=htmlencode2(request("mailtext"))
   title=request("title")
   if mailtext="" then 
   Response.End
   end if
   Set rs = Server.CreateObject("ADODB.Recordset")
   sql="select * from cmailbox"
   rs.open sql,conn,3,3
   rs.addnew
   rs("reid")=reid
   rs("senduid")=uname
   rs("sendname")=sendname
   rs("title")=title
   rs("mailtext")=mailtext
   rs("sdate")=now()
   rs.update
 response.write"<SCRIPT language=JavaScript>alert('信件发送成功！');"
 response.write"javascript:window.close();</script>" %>

