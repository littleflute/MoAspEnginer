<!--#include file="conn.asp"-->
<%
id = trim(request("id"))
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�������½��/ID');history.go(-1);</script>"
 response.end
end if
if IsNumeric(id) then
	sql = "select * from teacher where teacherid="&id
else
	sql = "select * from teacher where loginname='"&id&"'"
end if
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 response.write "<script>alert('û�иý�ʦ����ȷ��������ĵ�½��/ID�Ƿ���ȷ');history.go(-1);</script>"
else
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
<title>ȡ������</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="getpwdok.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="300" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>ȡ�����루�ڶ�����</td></tr>
  <tr>
    <td align=center>��½��/ID</td>
    <td>&nbsp;<%=id%><input type=hidden name="id" value="<%=rs("teacherid")%>"></td>
  </tr>
  <tr>
    <td align=center>����ȡ������</td>
    <td>&nbsp;<%=rs("ask")%></td>
  </tr>
  <tr>
    <td align=center>����ȡ�ش�</td>
    <td>&nbsp;<input type=text name="answer" size=15></td>
  </tr>
  <tr>
    <td align=center>������������</td>
    <td>&nbsp;<input type=password name="password" size=15></td>
  </tr>
  <tr>
    <td align=center>����һ��������</td>
    <td>&nbsp;<input type=password name="password1" size=15></td>
  </tr>
</table>
<center><input type=submit name="submit" value="ȷ��"></center>
</form>
</body>
</html>
<%
end if
rs.close
set rs = nothing
conn.close
set conn = nothing
%>