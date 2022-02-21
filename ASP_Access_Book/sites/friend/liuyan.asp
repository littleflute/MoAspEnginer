<%
Option Explicit

dim rs_lar,rs_word
dim sql
dim usid 
usid=session("user_id")
%>
<!--#include file="conn.asp"-->
<%
'叛断此用户是否已经注册

if session("user_id")=1 then response.redirect "notreg.htm"
Set rs_lar = Server.CreateObject("ADODB.Recordset")
sql="select * from larchives where user_id =" & session("user_id")
rs_lar.open sql,conn,3,2

'叛断此用户是否已经提交档案
if rs_lar.eof and rs_lar.bof then
	 response.redirect "notregist.htm"
	 response.end
end if

Set rs_word = Server.CreateObject("ADODB.Recordset")
'sql="select * from leaveword where for_id=" & session("my_fengyun_name") 
Sql="select * from leaveword where for_id="&usid&" order by id desc"
rs_word.open sql,conn,1,1
%>


<html>

<head>
<script language="JavaScript">
<!--
function MM_goToURL() { //v3.0
if (confirm("您确定要删除这条留言吗？按确定删除。")==1) {
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2)  eval(args[i]+".location='"+args[i+1]+"'");
}
}
//-->
</script>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>网友留言</title>
<style>
<!--
tr           { font-size: 12px; text-decoration: none}
body         { font-size: 12px; text-decoration: none}
a:link       { color: #003366; text-decoration: none ; font-size: 12px}
a:visited    { color: blue; text-decoration: none }
a:active     { color: red; text-decoration: none }
a:hover      { color: red; text-decoration: none }
td {  font-size: 12px; text-decoration: none}
-->
</style>
</head>

<body leftmargin="5" topmargin="0">
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
  </tr>
</table>
<table border="0" width="772" cellspacing="0" cellpadding="2">
  <tr> 
    <td colspan="3"><b><font color="#008000"> &nbsp;&nbsp;&nbsp;网友留言&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;共<font color="#ff0000"> 
      <%=rs_word.recordcount%></font> 条</font></b></td>
    <td width="414" align="right">[<a href="your.asp?user_id=<%=session("user_id")%>">交友请求</a>] 
      [<a href="#">我的留言</a>]&nbsp;[<a href="friend.asp?user_id=<%=session("user_id")%>">我的网友</a>]&nbsp;[<a href="read.asp?user_id=<%=session("user_id")%>">我的档案</a>]&nbsp;[<a href="edit.asp?user_id=<%=session("user_id")%>">修改档案</a>]&nbsp;[<a href="sendphoto.asp">相片管理</a>]&nbsp;</td>
  </tr>
  <tr> 
    <td colspan="4"><%do while not(rs_word.eof)%>  
      <table width="100%" border="0" cellspacing="1" cellpadding="5" bgcolor="#FF0033">
        <tr bgcolor="#FFD7EB"> 
          <td width="30%" height="22">时间：<%=rs_word("date")%></td>
          <td width="50%" height="22">&nbsp;留言者： <a href="read1.asp?user_id=<%=rs_word("user_id")%>"><font color="#ff0000"><%=rs_word("netname")%></font></a></td>
          <td width="20%" height="22" align="center"><a href="leaveword.asp?user_id=<%=rs_word("user_id")%>">[回复]</a>&nbsp;&nbsp;&nbsp;[<a href="#" onclick="MM_goToURL('parent','delliuyan.asp?id=<%=rs_word("id")%>');return document.MM_returnValue">删除</a>]</td>
        </tr>
        <tr bgcolor="#FFF0E1"> 
          <td colspan="3"><font color="#005716"><%=rs_word("word")%></font></td>
        </tr>
        
      </table>
      <br>
      <%rs_word.movenext%> <%loop%> </td>
  </tr>
</table>



      
<map name="Map">
  <area shape="rect" coords="232,95,275,111" href="default.asp">
  <area shape="rect" coords="293,93,361,113" href="your.asp">
  <area shape="rect" coords="379,93,445,113" href="list.asp">
  <area shape="rect" coords="470,93,532,114" href="register.asp">
  <area shape="rect" coords="552,94,620,114" href="sendphoto.asp">
  <area shape="rect" coords="640,93,702,114" href="pics.asp">
</map>
</body>
<%rs_lar.close:set rs_lar=nothing
 rs_word.close:set rs_word=nothing
 %>