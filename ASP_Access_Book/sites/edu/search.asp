<!--#include file="conn.asp"-->
<!--#include file="fenlei.asp"-->
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
<title>�ۺ�����</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="list.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 align=center cellPadding=5 width="450" border=1>
<tr><td>
<table width=100% border=0 align="center">
<tr><td align="center" class="header" colspan=2>�ۺ�����</td></tr>
  <tr>
    <td align="right">����<%=strfenlei1%>��</td>
    <td><input type=text name="fenlei1" size=15></td>
  </tr>
  <tr>
    <td align="right">����<%=strfenlei2%>��</td>
    <td><input type=text name="fenlei2" size=15></td>
  </tr>
  <tr>
    <td align="right">��ʦ������</td>
    <td><input type=text name="teacher" size=15></td>
  </tr>
  <tr>
    <td align="right">������ϣ�</td>
    <td><input type=text name="course" size=15></td>
  </tr>
  <tr>
    <td align="right">���ϱ��⣺</td>
    <td><input type=text name="title" size=15></td>
  </tr>
<%
sql = "select * from type"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
  <tr>
    <td align="right">�������ͣ�</td>
    <td>
        <select name="filetype">
        <option value="" selected>��ѡ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
<%
do while not rs.eof
%>
        <option value="<%=rs("typeid")%>"><%=rs("type")%></option>
<%
rs.movenext
loop
rs.close
set rs = nothing
%>
        </select>
    </td>
  </tr>
  <tr>
    <td align="center" colspan=2>ע��ֻ��д��Ҫ�����Ĳ��ּ��ɣ�����ȫ����д��֧��ģ������</td>
  </tr>
</table>
</td></tr></table>
<center><br>
<input type=submit name="submit" value="����">&nbsp;&nbsp;
<input type=reset name="reset" value="���">
</center>
</form>
<!--#include file="foot.asp"-->
</body>
</html>
<%
conn.close
set conn = nothing
%>