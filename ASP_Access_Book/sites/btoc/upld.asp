<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���������������վ�����ϴ��������</title>
</head>

<body>

<%
	if Request("T1") <> "Homestar" or Request("T2") <> "hd51988" then
%>
<div align="center">
  <center><form name="CheckAccount" action="upld.asp" method="POST">
  <table border="0" cellpadding="0" cellspacing="0" width="30%">
    <tr>
      <td width="100%" colspan="2">
        <p align="center"><font size="2"><b>����Ա��¼:</b></font></p>
      </td>
    </tr>
    <tr>
      <td width="31%" align="right"><b><font size="2">�˺�</font></b></td>
      <td width="69%"><input type="text" name="T1" size="20"></td>
    </tr>
    <tr>
      <td width="31%" align="right"><b><font size="2">����</font></b></td>
      <td width="69%"><input type="password" name="T2" size="20"></td>
    </tr>
    <tr>
      <td width="31%"></td>
      <td width="69%"><input type="submit" value="��¼" name="B1">&nbsp; <input type="reset" value="��д" name="B2"></td>
    </tr>
  </table></form>
  </center>
</div>

<%
	else
%>

<form name="me">
  <p align="center"><input type="button" value="�²�Ʒ���ݵ��ϴ�" name="B1" onClick="
  			me.action='sheet.asp?OP=����'
  			me.method='post'
  			me.submit()
  			"></p>
  <p align="center"><input type="button" value="���������Ʒ���" name="B2" onClick="
  			me.action='tag_list.asp'
  			me.method='post'
  			me.submit()
  			"></p>
  <p align="center"><input type="button" value="�������в�Ʒ����" name="B3" onClick="
  			me.action='pro_list.asp'
  			me.method='post'
  			me.submit()
  			"></p>
</form>

<%
	end if
%>
</body>

</html>
