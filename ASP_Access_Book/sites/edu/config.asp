<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
sql = "select * from config"
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
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
<title>ϵͳ����</title>
</head>
<body text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<br>
<form action="configok.asp" method="post">
<table style="BORDER-COLLAPSE: collapse" borderColor=#808080 width="400" border="1" align="center" cellpadding=1>
<tr><td align="center" class="header" colspan=2>ϵͳ����</td></tr>
  <tr>
    <td align="center">ѧУ����</td>
    <td>&nbsp;<input type=text name="schoolname" size=15 value=<%=rs("schoolname")%>></td>
  </tr>
  <tr>
    <td align="center">��Ȩ���к�</td>
    <td>&nbsp;<input type=text name="sn" size=15 value=<%=rs("sn")%>></td>
  </tr>
  <tr>
    <td align="center">����Ա����</td>
    <td>&nbsp;<input type=password name="adminpwd" size=15>����������ԭ���룩</td>
  </tr>
  <tr>
    <td align="center">������һ������</td>
    <td>&nbsp;<input type=password name="adminpwd1" size=15></td>
  </tr>
  <tr>
    <td align="center">һ��������</td>
    <td>&nbsp;<input type=text name="strfenlei1" size=15 value=<%=rs("strfenlei1")%>></td>
  </tr>
  <tr>
    <td align="center">����������</td>
    <td>&nbsp;<input type=text name="strfenlei2" size=15 value=<%=rs("strfenlei2")%>></td>
  </tr>
  <tr>
    <td align="center">����ԱE-Mail</td>
    <td>&nbsp;<input type=text name="adminmail" size=15 value=<%=rs("adminmail")%>></td>
  </tr>
  <tr>
    <td align="center">����Ա��ϵ��ʽ</td>
    <td>&nbsp;<input type=text name="contactadmin" size=15 value=<%=rs("contactadmin")%>></td>
  </tr>
  <tr>
    <td align="center">ϵͳURL</td>
    <td>&nbsp;<input type=text name="siteurl" size=15 value=<%=rs("siteurl")%>>������ַ��</td>
  </tr>
  <tr>
    <td align="center">ÿҳ��ʾ����</td>
    <td>&nbsp;<input type=text name="pagesize" size=15 value=<%=rs("pagesize")%>>�����б��ÿҳ��ʾ����</td>
  </tr>
  <tr>
    <td align="center">�Ƿ��������û�</td>
    <td>
<%
if rs("locked") = 1 then
 response.write "&nbsp;<input type=radio name=locked value='0'>������"
 response.write "<input type=radio name=locked value='1' checked>����"
else
 response.write "&nbsp;<input type=radio name=locked value='0' checked>������"
 response.write "<input type=radio name=locked value='1'>����"
end if
%>
    </td>
  </tr>
  <tr>
    <td align="center">�ϴ���ʽ</td>
    <td>
        &nbsp;<select name="upload_type">
        <option value=0<%if rs("upload_type")=0 then response.write(" selected")%>>������ϴ�</option>
        <option value=1<%if rs("upload_type")=1 then response.write(" selected")%>>Lyfupload</option>
        <option value=2<%if rs("upload_type")=2 then response.write(" selected")%>>Aspupload3.0</option>
        <option value=3<%if rs("upload_type")<>0 and rs("upload_type")<>1 and rs("upload_type")<>2 then response.write(" selected")%>>�ر��ϴ�����</option>
        </select>
    </td>
  </tr>
  <tr>
    <td align="center">�ϴ���С����</td>
    <td>&nbsp;<input type=text name="upload_size" size=15 value=<%=rs("upload_size")%>>K</td>
  </tr>
  <tr>
    <td align="center">��ֹ�ϴ�������</td>
    <td>&nbsp;<input type=text name="upload_filetype" size=40 value=<%=rs("upload_filetype")%>><br>
        &nbsp;����չ��֮�����Զ��ŷָ���
    </td>
  </tr>
  <tr>
    <td align="center">����Ա����</td>
    <td>&nbsp;<textarea name="gonggao" cols="35" rows="4"><%=rs("gonggao")%></textarea></td>
  </tr>
  <tr>
    <td align="center">ҳ�涥��<br>������</td>
    <td>&nbsp;<textarea name="headads" cols="35" rows="4"><%=rs("headads")%></textarea></td>
  </tr>
</table>
<center><input type=submit name="submit" value="ȷ��">&nbsp;&nbsp;<input type=reset name="reset" value="����"></center>
</form>
</body>
</html>
<%
rs.close
set rs = nothing
conn.close
set conn = nothing
%>