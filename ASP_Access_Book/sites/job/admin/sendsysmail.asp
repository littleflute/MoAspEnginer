<% Response.Buffer=True %>
<!-- #include file="../inc/admin.inc" -->
<!--#include file="../inc/html.inc"-->
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
<form method=post action="sendsysmail.asp?stype=<%=request("stype")%>">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="408" height="231" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
      <td height="15" width="406" valign="bottom" background="../images/t-bg1.gif">
        <p align="center">=== 群发站内信件 ===</td>       
    </tr>
	<tr>
        <td height="216" width="406" valign="top" bgcolor="#fffff4"> 
          <p align="center"><br>
        标 题:        
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
</FORM>
</body>
</html>
<% mailtext=htmlencode2(request("mailtext"))
   title=request("title")
   if mailtext="" then 
   Response.End
   end if
   if title="" then title="无标题" end if
   senduid="sysop"
   sendname="系统管理员"
   
   if request("stype")="person" then
   Set rs = Server.CreateObject("ADODB.Recordset")
   sql="select uname from person"
   rs.open sql,conn,1,1
   i=0
   do while not rs.eof
   who=rs("uname")
   SQL="insert into pmailbox(reid,senduid,sendname,title,sdate,mailtext,newmail) values('" & who & "','" & senduid & "','" & sendname & "','" & title & "',now(),'" & mailtext & "',0)" 
   conn.Execute SQL
   rs.MoveNext 
   i=i+1
   loop 
   rs.close 
   set rs=nothing
   set conn=nothing 
   mailnum=i
   end if
   
   if request("stype")="company" then
   Set rs = Server.CreateObject("ADODB.Recordset")
   sql="select uname from company"
   rs.open sql,conn,1,1
   i=0
   do while not rs.eof
   who=rs("uname")
   SQL="insert into cmailbox(reid,senduid,sendname,title,sdate,mailtext,newmail) values('" & who & "','" & senduid & "','" & sendname & "','" & title & "',now(),'" & mailtext & "',0)" 
   conn.Execute SQL
   rs.MoveNext 
   i=i+1
   loop 
   rs.close 
   set rs=nothing
   set conn=nothing
   mailnum=i
   end if
response.write"<SCRIPT language=JavaScript>alert('信件发送成功,共发送"&mailnum&"封站内信件！');"
response.write"javascript:window.close();</SCRIPT>" %>
