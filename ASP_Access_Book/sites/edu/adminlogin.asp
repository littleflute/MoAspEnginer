<!--#include file="conn.asp"-->
<%
dim num1
dim rndnum
Randomize
Do While Len(rndnum)<4
 num1=CStr(Chr((57-48)*rnd+48))
 rndnum=rndnum&num1
loop
session("verifycode")=rndnum
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE>TD {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
BODY {
	FONT-SIZE: 9pt; LINE-HEIGHT: 140%
}
A:link {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:visited {
	COLOR: #0033cc; TEXT-DECORATION: none
}
A:active {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:hover {
	COLOR: #000000; TEXT-DECORATION: underline
}
.header	{
font-family: Tahoma, Verdana; font-size: 9pt; color: #FFFFFF; background-color: #00CCCC
}
.category{
font-family: Tahoma, Verdana; font-size: 9pt; color: #000000; background-color: #EFEFEF
}
</STYLE>
<title>管理员登陆</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="admincheck.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="300" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>管理员登陆</td></tr>
  <tr>
    <td align="center">请输入管理员密码</td>
    <td>&nbsp;<input type=password name="adminpwd" size=15></td>
  </tr>
<tr><td align="center">请输入右侧的验证码</td><td>&nbsp;<input type=text name="verifycode" size=15>
<img src="num.ASP">
</td></tr>
</table>
<center><br><input type=submit name="submit" value="确认">&nbsp;&nbsp;<input type=reset name="reset" value="清空"></center>
</form>
</body>
</html>
<%
conn.close
set conn = nothing
%>