<!--#include file="conn.asp"-->
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
<title>取回密码</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="getpwd2.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="250" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>取回密码（第一步）</td></tr>
  <tr>
    <td align=center>登陆名/ID</td>
    <td>&nbsp;<input type=text name="id" size=15 title="请输入您的登陆名或ID"></td>
  </tr>
</table>
<center><input type=submit name="submit" value="下一步"></center>
</form>
</body>
</html>