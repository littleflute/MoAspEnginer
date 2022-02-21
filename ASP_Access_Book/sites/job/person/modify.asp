<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid")
if request("del")="true" then
conn.Execute("delete * from person where uname='"&uname&"'")
conn.Execute("delete * from pmailbox where reid='"&uname&"'")
conn.Execute("delete * from pfavorite where uname='"&uname&"'")
conn.Execute("delete * from cfavorite where fuid='"&uname&"'")
response.write"<SCRIPT language=JavaScript>alert('用户注销成功！');"
response.write"this.location.href='../';</SCRIPT>"
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'and job<>'""'"
rs.open sql,conn,1,1 %>
<%  if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('您尚未登录求职简历，请先登录求职简历！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;首页</title>
</head>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="149">
    <tr>
      <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        　 </td>
    </tr>
    <tr>
      <td width="139" valign="top" bgcolor="#53BEB0"> 　 
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top" align="center">
        <p align="center"><br>
        <a href="main.asp">登录首页</a><br>
        <br>
        <a href="register.asp">登录求职简历</a><br>
        <br>
        <a href="modify.asp">更新求职简历</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">修改登录密码</a><br>
        <br><a href="search.asp">全部职位列表</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top" align="center">
        <p align="center"><br>
        <br><a href="favorite.asp">我的收藏夹</a><br>
        <br>
        <a href="mailbox.asp">我的信箱</a><br>
        <br>
        <a href="../exit.asp">退出登录</a>
              </td>
            </tr>
          </table>
          </center>
        </div>
        <p align="center">&nbsp; </td>
      <td width="38" height="413" valign="top"><img border="0" src="../images/selfk.GIF"></td>
      <td width="455" height="413" valign="top">
  </center>
    
      <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="241" height="195">
          <tr>
            <td width="239" height="148">
      <p align="left"><br>
      <br>
      您需要更新求职简历的哪一部分？<br>
      <a href="register.asp?modify=ture"><br>
      </a><img border="0" src="../images/push.gif"> <a href="register.asp?modify=ture">个人基本资料</a></p>   
      <p align="left"><img border="0" src="../images/push.gif"> <a href="register2.asp?modify=ture">个人主要特长—相关工作经历</a> </p>    
      <p align="left"><img border="0" src="../images/push.gif"> <a href="register3.asp?modify=ture">希望工作条件—联系信息<br>    
      <br>
      </a></p>
            </td>
          </tr>
          <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
          </tr>
          <tr>
            <td width="239" height="48">
      <p align="left"><img border="0" src="../images/push.gif"> <a href="#"onclick="javascript:if (confirm('是否确定要注销帐号?')) href='modify.asp?del=true'; 
	    else return;">注销帐号</a>
            </td>
          </tr>
        </table>
      </div>
    
      <p>　
    
  </td>
  <center>
      <td width="1" height="413" valign="top" bgcolor="#00006A"></td>
      <td width="133" height="413" valign="top" bgcolor="#F3F3F3">　</td>
    </tr>
    <tr>
      <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</body>

</html>
