<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
id = trim(request("id"))
returnlist = trim(request("returnlist"))
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
<title>删除教师</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<form action="delteacherok.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="250" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header">将有下列数据被删除</td></tr>
<tr><td align="left">
1.该教师的个人资料<br>
2.该教师发布的所有资料<br>
3.该教师上传的所有资料
</td></tr>
</table>
<center><input type=hidden name="id" value="<%=id%>"><input type=hidden name="returnlist" value="<%=returnlist%>"><br>
<input type=submit name="submit" value="确定">&nbsp;&nbsp;<input type=button name="cancle" value="取消" onclick="history.go(-1);"></center>
</form>
</body>
</html>
<%
conn.close
set conn = nothing
%>