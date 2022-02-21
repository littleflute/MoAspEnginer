<!--#include file="conn.asp"-->
<%
dim rs_lar,rs_word,rs_apply,rs_back,rs_friend,rs_user
dim sql

'叛断Session变量是否超时
if isempty(session("user_id")) then
   response.redirect "timeout.htm"
end if

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
sql="select * from leaveword where for_id=" & session("user_id") & " order by id desc"
rs_word.open sql,conn,1,1

set rs_apply=server.createobject("adodb.recordset")
sql="select * from apply where for_id=" & session("user_id")
rs_apply.open sql,conn,3,2

set rs_back=server.createobject("adodb.recordset")
sql="select * from back where for_id=" & session("user_id")
rs_back.open sql,conn,3,2

set rs_friend=server.createobject("adodb.recordset")
sql="select * from friend where for_id=" & session("user_id")
rs_friend.open sql,conn,3,2

Set rs_user = Server.CreateObject("ADODB.Recordset")
sql="select * from user_reg where user_id=" & session("user_id")
rs_user.open sql,conn,3,2

%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Namo WebEditor v4.0(Trial)">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>您的入会日期是</title>
<style>
<!--
.x:link      { color: white; text-decoration: none }
.x:visited   { color: white; text-decoration: none }
.x:active    { color: red; text-decoration: none }
.x:hover     { color: red;text-decoration:nome}
tr           { font-size: 10pt }
body         { font-size: 10pt }
a:link       { color: blue; text-decoration: none }
a:visited    { color: blue; text-decoration: none }
a:active     { color: red; text-decoration: none }
a:hover      { color: red; text-decoration: none }
-->
</style>
<style type="text/css">
<!--
body {
	background-image: url(bihaibg.jpg);
}
-->
</style>
</head>

<body leftmargin="5" topmargin="0">
<div align="center">
  <table width="772" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td><img src="images/top.gif" width="772" height="112" usemap="#Map" border="0"></td>
    </tr>
  </table>
  <table border="0" width="772" height="176" cellspacing="0" cellpadding="0">
    <tr>
      
      <td width="507" height="170" valign="top"> 
        <table border="0" width="100%" cellspacing="0" cellpadding="0" height="16">
          <tr bgcolor="#CDDCEB"> 
            <td colspan="7" width="424" height="18"> 有如下网友向您提出交友请求:</td>
            </tr>
          <tr>
            <td width="76" align="center" height="1" style="border-bottom-style: solid">日期</td>
            <td width="51" align="center" height="1" style="border-bottom-style: solid">姓名</td>
            <td width="41" align="center" height="1" style="border-bottom-style: solid">性别</td>
            <td width="42" align="center" height="1" style="border-bottom-style: solid">年龄</td>
            <td width="71" align="center" height="1" style="border-bottom-style: solid">来自</td>
            <td width="65" align="center" height="1" style="border-bottom-style: solid">网名</td>
            <td width="80" align="center" height="1" style="border-bottom-style: solid">态度</td>
            </tr>
          <%do while not(rs_apply.eof)%>
          
          <tr bgcolor="#FFEBE1"> 
            <td width="76" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_apply("apply_date")%></td>
              
          <td width="51" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><a href="read.asp?user_id=<%=rs_apply("user_id")%>"><%=rs_apply("user_name")%></a></td>
              
          <td width="41" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_apply("user_sex")%></td>
              
          <td width="42" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_apply("user_age")%></td>
              
          <td width="71" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_apply("user_home")%></td>
              
          <td width="65" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_apply("user_netname")%></td>
              
          <td width="80" height="1" align="center" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><a href="accept.asp?id=<%=rs_apply("id")%>&user_id=<%=rs_apply("user_id")%>">[愿意]</a><a href="refuse.asp?id=<%=rs_apply("id")%>&user_id=<%=rs_apply("user_id")%>">[拒绝]</a></td>
            </tr>
          <%rs_apply.movenext%>
          <%loop%>
          <tr>
            <td width="125" colspan="2" height="15">&nbsp;&nbsp;&nbsp;</td>
            <td width="81" colspan="2" height="15"></td>
            <td width="69" height="15"></td>
            <td width="63" height="15"></td>
            <td width="78" height="15"></td>
            </tr>
          </table>
          
      <table border="0" width="100%" cellspacing="0" cellpadding="2">
        <tr bgcolor="#CDDCEB"> 
          <td colspan="5" height="18"> 您的交友请求有如下回复:</td>
          </tr>
        <tr bgcolor="#F2F2F2"> 
          <td width="21%" align="center" style="border-bottom-style: solid">日期</td>
            <td width="16%" align="center" style="border-bottom-style: solid">姓名</td>
            <td width="9%" align="center" style="border-bottom-style: solid">性别</td>
            <td width="13%" align="center" style="border-bottom-style: solid">网名</td>
            <td width="38%" align="center" style="border-bottom-style: solid">　</td>
          </tr>
        <%do while not(rs_back.eof)%> 
        <tr bgcolor="#FFF2EC"> 
          <td width="21%" align="center"><%=rs_back("back_date")%></td>
            <td width="16%" align="center"><%=rs_back("name")%></td>
            <td width="9%" align="center"><%=rs_back("sex")%></td>
            <td width="13%" align="center"><%=rs_back("netname")%></td>
            <td width="38%" align="center"> <%if rs_back("result")="已接受" then%> 
              已接受 <a href="moveto.asp?id=<%=rs_back("id")%>">[移入好友列表]</a> <%end if%> 
              <%if rs_back("result")="已拒绝" then%> 已拒绝 <%end if%> [<a href="delqq.asp?id=<%=rs_back("id")%>">删除</a>]</td>
          </tr>
        <%rs_back.movenext%> <%loop%> 
        <tr> 
          <td width="21%">&nbsp;</td>
            <td width="16%"></td>
            <td width="9%"></td>
            <td width="13%"></td>
            <td width="38%"> </td>
          </tr>
        </table>
  
        
        <table border="0" width="100%" cellspacing="1" cellpadding="2">
          <tr bgcolor="#EBEBEB"> 
            <td colspan="4"><b><font color="#008000"><u> 网友留言前5条 共<%=rs_word.recordcount%>条 
              <a href="liuyan.asp">查看全部留言</a> </u></font></b></td>
            </tr>
          <tr>
            
            <td width="90" align="center" style="border-bottom-style: solid">日期</td>
              
          <td width="61" align="center" style="border-bottom-style: solid">网名</td>
              
          <td width="261" align="center" style="border-bottom-style: solid" height="18">内容</td>
            <td width="98" align="center" style="border-bottom-style: solid">　</td>
            </tr>
          <%do while not(rs_word.eof)%>
          
          <tr> 
            <td width="90" align="center" bgcolor="#FFEBE1" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="23"><%=rs_word("date")%></td>
              
          <td width="61" align="center" bgcolor="#FFEBE1" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="23"><a href="read1.asp?user_id=<%=rs_word("user_id")%>"><%=rs_word("netname")%></a></td>
              
          <td width="261" height="23" bgcolor="#FFEBE1" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF"><%=rs_word("word")%></td>
              
          <td width="98" align="center" bgcolor="#FFEBE1" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="23"><a href="leaveword.asp?user_id=<%=rs_word("user_id")%>">[回复]</a>[<a href="delliuyan.asp?id=<%=rs_word("id")%>">删除</a>]</td>
            </tr>
          <%rs_word.movenext%>
          <%loop%>
          <tr>
            
            <td width="90" align="center">&nbsp;</td>
            <td width="61" align="center"></td>
              
          <td width="261" align="center"></td>
              
          <td width="98" align="center"></td>
            </tr>
          </table>
  

  
      <table border="0" height="2" cellspacing="0" cellpadding="0" width="100%">
        <tr>
          <td height="6" width="498" style="border-bottom-style: solid" bordercolor="#000000">&nbsp;&nbsp;&nbsp;</td>
    </tr>
        
        <tr bgcolor="#FFF5F0"> 
          <td height="17" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF">您的人气度为<font color="red"><%=rs_lar("renqi")%></font></td>
    </tr>
        <tr bgcolor="#FFF5F0"> 
          <td height="17" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF">您的档案已被浏览了<font color="red"><%=rs_lar("click")%></font>次</td>
    </tr>
        <tr bgcolor="#FFF5F0"> 
          <td height="17" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF">您的入会日期是:<%=formatdatetime(rs_user("date"),vblongdate)%></td>
    </tr>
        <tr bgcolor="#FFF5F0"> 
          <td height="17" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF">您的最后修改日期是:<%=formatdatetime(rs_lar("date"),vblongdate)%></td>
    </tr>
        <tr bgcolor="#FFF5F0"> 
          <td height="17" style="border-bottom-style: solid" bordercolor="#000000"> 
            <a href="read.asp?user_id=<%=session("user_id")%>">[查看档案]</a><a href="edit.asp?user_id=<%=session("user_id")%>">[修改档案]</a><a href="sendphoto.asp">[相片管理]</a></td>
    </tr>
  </table>      </td>
        
    <td width="5" height="170" valign="top"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
        
    <td width="259" height="170" valign="top"> 
      <table border="0" width="100%" cellspacing="0" cellpadding="1">
        <tr> 
          <td height="18" colspan="4"><font color="#008000"><b><u>您的网友列表</u></b></font></td>
          </tr>
        <tr bgcolor="#006600"> 
          <td width="66" height="18" align="center" bgcolor="#CDDCEB"><font color="#000000">网名</font></td>
            <td width="30" height="18" align="center" bgcolor="#CDDCEB"><font color="#000000">性别</font></td>
            <td width="115" height="18" align="center" bgcolor="#CDDCEB"><font color="#000000">来自</font></td>
            <td width="33" height="18" align="right" bgcolor="#CDDCEB">&nbsp;</td>
          </tr>
        <%do while not(rs_friend.eof)%> 
        <tr> 
          <td width="66" bgcolor="#EEEEEE" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="18"><a href="read1.asp?user_id=<%=rs_friend("user_id")%>"><%=rs_friend("netname")%></a></td>
            <td width="30" align="center" bgcolor="#EEEEEE" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="18"><%=rs_friend("sex")%></td>
            <td width="115" bgcolor="#EEEEEE" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="18"><%=rs_friend("home")%></td>
            <td width="33" bgcolor="#EEEEEE" style="border-bottom-style: solid; border-bottom-width: 1" bordercolor="#FFFFFF" height="18" align="right"><a href="delhy.asp?id=<%=rs_friend("id")%>">删除</a></td>
          </tr>
        <%rs_friend.movenext%> <%loop%> 
        </table>
  

  
      <br>
      <a href="pics.asp" target="_blank"><font color="#660033" size="2">网友照片</font></a>      </td>
      </tr>
    </table>
  <map name="FPMap1"> 
    <area shape="rect" coords="381, 4, 446, 18" href="sendphoto.asp"> 
    <area shape="rect" coords="283, 4, 357, 18" href="register.asp"
 target="_blank"> 
    <area shape="rect" coords="190, 4, 264, 18" href="list.asp"> 
    <area shape="rect" coords="8, 2, 71, 16" href="your.asp"> 
    <area shape="rect" coords="102, 4, 168, 18" href="no.htm"> 
  </map>
  <map name="Map"> 
    <area shape="rect" coords="231,93,279,111" href="default.asp">
    <area shape="rect" coords="289,93,358,111" href="your.asp">
    <area shape="rect" coords="379,94,448,112" href="list.asp">
    <area shape="rect" coords="467,93,532,109" href="register.asp">
    <area shape="rect" coords="553,93,620,113" href="sendphoto.asp">
    <area shape="rect" coords="640,93,705,111" href="pics.asp">
  </map>
</div>
</body>
<%rs_lar.close:set rs_lar=nothing
 rs_word.close:set rs_word=nothing
 rs_apply.close:set rs_apply=nothing
 rs_friend.close:set rs_friend=nothing
 rs_user.close:set rs_user=nothing%>