<%
if session("user_id")<>1 then
   response.redirect "havereg.htm"
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>交友申请</title>
<style>
<!--
-->
</style>
<script language="javascript">
<!--
function isok(theform)
{
if (theform.user_name.value.length<2 || theform.user_name.value.length>10)
  {
    alert("用户名称最少要2位，最多10位！");
    theform.user_name.focus();
    return (false);
  }
if (theform.password.value.length<3 || theform.password.value.length>10)
  {
    alert("密码最少要3位，最多10位！ ！");
    theform.password.focus();
    return (false);
  }
if (theform.password_two.value=="")
  {
    alert("你密还有确认密码没填呢！");
    theform.password_two.focus();
    return (false);
  }

if (theform.password_two.value!=theform.password.value)
  {
    alert("两次密码怎么不一样！");
    theform.password_two.focus();
    return (false);
  }

return (true);
}
-->
</script>
<link rel="stylesheet" href="ysb.css">
</head>

<body>
<form method="POST" action="regsubmit.asp" onsubmit="return isok(this)">
<table border="0" width="402" cellspacing="0" cellpadding="0" align="center">
    <tr align="center"> 
      <td width="400" height="25"><b>网友入会注册</b></td>
  </tr>
  <tr>
      <td width="400" height="20"><font color="#FF0000">请填好下表提交，并记住用户名和密码，用于以后登陆时的</font><font color="#008000"><a href="#">身份验证</a></font></td>
  </tr>
</table>
  <table border="1" width="400" bordercolor="#CC0000" cellspacing="0" cellpadding="0" height="66" align="center" bgcolor="#067B0F">
    <tr>
      <td width="100%" valign="top" height="43">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
          <tr>
            <td width="43%" height="30"> 
              <p align="center">用户名称:</td>
            <td width="57%" height="30"> 
              <input type="text" name="user_name" size="18">(2-10)</td>
          </tr>
          <tr>
            <td width="43%" height="30"> 
              <p align="center">管理密码:</td>
            <td width="57%" height="30">
<input type="password" name="password" size="18">(3-10)</td>
          </tr>
          <tr>
            <td width="43%" height="30"> 
              <p align="center">确定密码:</td>
            <td width="57%" height="30">
<input type="password" name="password_two" size="18"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<p align="center">　<input type="submit" value="提交" name="B1"><input type="reset" value="全部重写" name="B2"></p>
</form>

</body>

</html>
