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
<title>��ʦע��</title>
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
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<!--#include file="head.asp"-->
<form action="registerok.asp" method="post" onSubmit=submitonce(this)>
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="400" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>��ʦע��</td></tr>
  <tr>
    <td width=130>&nbsp;��������������<%=strfenlei1%></td>
    <td width=270>&nbsp;<input type=text name="fenlei1" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;��������������<%=strfenlei2%></td>
    <td>&nbsp;<input type=text name="fenlei2" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;��������</td>
    <td>&nbsp;<input type=text name="teacher" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;���ĵ�¼��</td>
    <td>&nbsp;<input type=text name="loginname" size=22>������ȫΪ���֣�</td>
  </tr>
  <tr>
    <td>&nbsp;����������</td>
    <td>&nbsp;<input type=password name="password" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;����һ������</td>
    <td>&nbsp;<input type=password name="password1" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;�����һ�����</td>
    <td>&nbsp;<input type=text name="ask" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;�����һش�</td>
    <td>&nbsp;<input type=text name="answer" size=25></td>
  </tr>
  <tr>
    <td>&nbsp;E-Mail</td>
    <td>&nbsp;<input type=text name="email" size=25>�����Բ��</td>
  </tr>
  <tr>
    <td>&nbsp;������ҳ</td>
    <td>&nbsp;<input type=text name="homepage" size=25>�����Բ��</td>
  </tr>
  <tr>
    <td>&nbsp;QQ����</td>
    <td>&nbsp;<input type=text name="qq" size=25>�����Բ��</td>
  </tr>
  <tr>
    <td>&nbsp;ͨѶ��ַ</td>
    <td>&nbsp;<input type=text name="address" size=25>�����Բ��</td>
  </tr>
  <tr>
    <td>&nbsp;���˼��</td>
    <td>
        &nbsp;<textarea name="intro" cols="25" rows="6"></textarea>
    </td>
  </tr>
</table>
<center><br><input type=submit name="submit" value="ȷ��">&nbsp;&nbsp;<input type=reset name="reset" value="���"></center>
</form>
<!--#include file="foot.asp"-->
</body>
</html>
<%
conn.close
set conn = nothing
%>