<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
<!--#include file="isadmin.asp"-->
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
<title>教师管理</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<form action="adminsearchteacher.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="300" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>教师管理</td></tr>
<tr><td align="center" colspan=2>请搜索要管理的教师，为了加快搜索速度，<br>请至少填一项，全部不填则显示所有教师清单</td></tr>
  <tr>
    <td width=100>&nbsp;&nbsp;&nbsp;教师所属<%=strfenlei1%></td>
    <td>&nbsp;<input type=text name="fenlei1" size=15></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;&nbsp;教师所属<%=strfenlei2%></td>
    <td>&nbsp;<input type=text name="fenlei2" size=15></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;&nbsp;教师姓名</td>
    <td>&nbsp;<input type=text name="teacher" size=15></td>
  </tr>
  <tr>
    <td>&nbsp;&nbsp;&nbsp;教师ID</td>
    <td>&nbsp;<input type=text name="id" size=15></td>
  </tr>
</table>
<center><input type=submit name="submit" value="搜索"></center>
</form>
</body>
</html>
<%
conn.close
set conn = nothing
%>