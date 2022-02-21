<% if session("puid")<>"" then uname=session("puid") end if
if session("cuid")<>"" then uname=session("cuid") end if
stype=request("stype") %>
<!-- #include file="inc/dbconn.inc" -->
<!-- #include file="inc/enpasswd.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>天天人才—&gt;修改密码</title>
</head>
<SCRIPT language=JavaScript>
<!--
function checkform()
{
 if (document.changepwd.pwd.value=="") {
		alert("请输入原密码！");
		changepwd.pwd.focus();		
		    return (false);
  }
 if (document.changepwd.newpwd.value=="") {
		alert("请输入新密码！");
		changepwd.pwd.focus();		
		    return (false);
  }
  if (document.changepwd.newpwd.value.length<3) {
    alert("密码不能少于三位！");
    return false;   
  }
  if (document.changepwd.newpwd.value != document.changepwd.newpwd2.value) {
    alert("两次密码输入不一致！");
    document.changepwd.newpwd.value="";
    document.changepwd.newpwd2.value="";
	return false;
  }
  return true;
}
//-->
</SCRIPT>
<form name=changepwd onsubmit="return checkform();" action="changepwd.asp?stype=<%=stype%>" method=post>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="256" height="180" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#EBEBEB">
    <tr>
      <td width="254" height="1" valign="top" background="images/t-bg1.gif">
        <p align="center">=== 修改密码 ===</td>  
    </tr>
    <tr>
        <td width="254" height="187" valign="top" bgcolor="#fffff4"> 
          <p align="center"><br>
        用户名:<font color="#0000B5"><%=uname%></font>&nbsp;[<%if stype="company" then response.write"单位用户" else response.write"个人用户" end if%>]</p>
        <p align="center">&nbsp;原密码:<input type="password" name="pwd" size="20" maxLength=20></p>
        <p align="center">&nbsp;新密码:<input type="password" name="newpwd" size="20" maxLength=20></p>
  </center>
        <p align="left">&nbsp;&nbsp; 重复新密码:<input type="password" name="newpwd2" size="20" maxLength=20></p>
  <center>
          <p align="center"><input type="submit" value="修 改" name="B1"><br>
        <br>
      </td>
    </tr>
  </table>
  </center>
</div>
</form>
<p align="center">【<a href="javascript:window.close()">关闭窗口</a>】</p>
<% pwd=mistake(request("pwd"))
if pwd="" then
   response.end
   end if
   newpwd=mistake(request("newpwd"))
   if stype="person" then
   set rs=server.createobject("adodb.recordset")
   sql1="select * from person where uname='"&uname&"' and pwd='"&pwd&"'"
   rs.open sql1,conn,3,3
   if rs.bof or rs.eof then
   response.write"<SCRIPT language=JavaScript>alert('原密码错误，请重新输入！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   else
   rs("pwd")=newpwd
   rs.update
   rs.close
   response.write"<SCRIPT language=JavaScript>alert('密码修改成功');"
   response.write"javascript:window.close();</SCRIPT>"
   end if
   end if
   
   if stype="company" then
   set rs=server.createobject("adodb.recordset")
   sql2="select * from company where uname='"&uname&"' and pwd='"&pwd&"'"
   rs.open sql2,conn,3,3
   if rs.bof or rs.eof then
   response.write"<SCRIPT language=JavaScript>alert('原密码错误，请重新输入！');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   else
   rs("pwd")=newpwd
   rs.update
   rs.close
   response.write"<SCRIPT language=JavaScript>alert('密码修改成功');"
   response.write"javascript:window.close();</SCRIPT>"
   end if 
   end if %>