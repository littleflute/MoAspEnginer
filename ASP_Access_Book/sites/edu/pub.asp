<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
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
<title>��������</title>
<script language="javascript">
function submitonce(theform)
{
//if IE 4+ or NS 6+
if (document.all||document.getElementById){
//screen thru every element in the form, and hunt down "submit" and "reset"
for (i=0;i<theform.length;i++){
var tempobj=theform.elements[i]
if(tempobj.type.toLowerCase()=="submit"||tempobj.type.toLowerCase()=="reset")
//disable em
tempobj.disabled=true
}
}
}
</script>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0 >
<br>
<form action="pubok.asp" method="post" onSubmit=submitonce(this)>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="480" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>��������</td></tr>

  <tr>
    <td align="center">��ʦ����</td>
    <td>&nbsp;<input type=text name="teacher" size=15 value="<%=teacher%>"></td>
  </tr>
  <tr>
    <td align="center">�γ�����</td>
    <td>&nbsp;<input type=text name="course" size=15>��Ϊ������ң������ʦע�⾡��ͳһ���ƣ�</td>
  </tr>
  <tr>
    <td align="center">���ϱ���</td>
    <td>&nbsp;<input type=text name="title" size=52></td>
  </tr>
  <tr>
    <td align="center">�����ϴ�</td>
    <td>&nbsp;<iframe name="up" frameborder=0 width=400 height=26 scrolling=no src=upload.asp></iframe></td>
  </tr>
  <tr>
    <td align="center">���ϵ�ַ</td>
    <td>&nbsp;<input type=text name="fileurl" size=52>
        <br>&nbsp;��������������������վ����������д���ӵ�ַ��
    </td>
  </tr>
  <tr>
    <td align="center">���ϴ�С</td>
    <td>&nbsp;<input type=text name="filesize" size=15>K</td>
  </tr>
  <tr>
    <td align="center">��������</td>
    <td>
<%
sql = "select * from type"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
%>
&nbsp;<select name="type">
<option value="" selected>��ѡ��</option>
<%
do while not rs.eof
%>
 <option value="<%=rs("typeid")%>"><%=rs("type")%></option>
<%
 rs.movenext
loop
response.write "</select>"
rs.close
set rs = nothing
%>
    </td>
  </tr>
  <tr>
    <td align="center">���ϼ��</td>
    <td>&nbsp;<textarea name="content" cols="51" rows="6"></textarea></td>
  </tr>
</table>
<center><br><input type=submit name="submit" value="����">&nbsp;&nbsp;<input type=reset name="reset" value="���"></center>
</form>
</body>
</html>
<%
conn.close
set conn = nothing
%>