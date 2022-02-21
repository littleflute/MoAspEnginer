<!--#include file="articleCHAR.INC"-->
<!--#include file="conn.asp"-->
<%
On Error Resume Next
dim JMail,body
dim rs_apply,rs_lar,rs_email,rs_word
dim user_id,netname,sql,Email,myEmail,myname,word,errst
'叛断Session变量是否超时
if isempty(session("user_id")) then
   response.redirect "timeout.htm"
end if

if session("user_id")=1 then
	response.redirect "notreg.htm"
	response.end
end if

user_id=request("user_id")
netname=request("netname")

Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id =" & session("user_id")
rs_lar.open sql,conn,3,2

netname        =rs_lar("netname")
word           =htmlencode2(request("word"))
user_id        =request("user_id")

if word<>"" then
if len(word)>1000 or word="" then
%> 
<link rel="stylesheet" href="card/ysb.css">

<table border="1" width="31%" bordercolor="#CC0000" cellspacing="0" align="center" bgcolor="#067B0F">
    <td> 
      <p align="center"><b><font color="#FFFFFF" size="2">错误提示</font></b> 
    <tr>
<td>
      <table width="100%" cellspacing="0" cellpadding="3">
        <tr bgcolor="#FFEBE1"> 
          <td width="100%"> 
            <p align="center"><font size="2">留言最多不能超过1000个符或不能为空!</font> 
          </td>
  </tr>
        <tr bgcolor="#FFEBE1"> 
          <td width="100%"> 
            <p align="center"><a href="" onclick="history.go(-1)"><font size="2">[返回]</font></a>
    </td>
  </tr>
</table>
</td>
</tr>
</table>

<%
response.end
end if
   Set rs_word = Server.CreateObject("ADODB.Recordset")
   rs_word.open "leaveword",conn,3,2
   rs_word.addnew
   rs_word("netname")=netname
   rs_word("word")=word
   rs_word("date")=date
   rs_word("for_id")=user_id
   rs_word("user_id")=session("user_id")
   rs_word.update
   rs_word.close
   set rs_word=nothing
   response.write "<script language='javascript'>" & vbcrlf
   response.write "history.go(-2);"
   response.write "</script>"

end if%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>签写留言</title>
<style>
<!--
tr           { font-size: 10pt }
body         { font-size: 10pt }
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: #ff9966; text-decoration: none }
a:hover      { color: red; text-decoration: none }
-->
</style>

</head>

<body>
<form>
<input type="hidden" name="user_id" value="<%=user_id%>">
<div align="center">
    <table border="0" cellspacing="1" height="111" cellpadding="0" bgcolor="#000000" width="95%">
      <tr bgcolor="#990000"> 
        <td height="25"> 
          <p align="center"><b><font color="#FFFFFF">写留言</font></b> <font color="#FFFF00">[不超过1000个字符]</font> 
        </td>
    </tr>
      <tr bgcolor="#FFFFCC"> 
        <td height="90"> 
          <table border="0" cellspacing="0" height="45" bgcolor="#FFEBE1" cellpadding="5" width="100%">
            <tr> 
              <td height="25">您的网名：</td>
              <td height="25"> 
                <input type="text" name="netname" size="20" value="<%=netname%>" disabled>
              </td>
            </tr>
            <tr> 
              <td height="1" valign="top" colspan="2">
                <hr size="1" noshade>
              </td>
            </tr>
            <tr> 
              <td valign="top">留言内容：</td>
              <td valign="top"> 
                <textarea rows="12" name="word" cols="80" wrap="hard"></textarea>
              </td>
            </tr>
            <tr> 
              <td colspan="2" valign="middle" align="center"> 
                <p> 
                  <input type="submit" value="提交" name="B1">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="reset" value="重写" name="B2">
              </td>
            </tr>
          </table>
      </td>
    </tr>
  </table>
</div>
</form>

</body>

</html>
