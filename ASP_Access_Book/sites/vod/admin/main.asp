<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<html>

<head>
<link rel="stylesheet" href="style.css">
</head>

<BODY bgcolor="#009458">
<P>&nbsp;</P>
<table border="0" cellspacing="1" width="80%" align=center bgcolor="#a5d0dc">
  <tr> 
    <td bgcolor="#FFFFFF"><img src="../b.jpg" width="760" height="103"></td>
  </tr>
  <tr>
    <td bgcolor="#A9E4D5"><p align=center><b>��Ƶ����ϵͳ����ҳ��</b></p>
      <p><b>����Ա�ɽ��в���˵��</b>��<br>
        <br>
        1�����Ѿ����ӰƬ�޸Ļ�ɾ����������������ӽ��в����������û��������û�����ͨ����Ա <br>
        <br>
        2����ӰƬ��Ŀ��������޸�ɾ����������������ӽ��в����������û��������û� <br>
        <br>
        3���û��������޸�ɾ����������������ӽ��в����������û��������û� <br>
        <br>
        4��<font color=red>Ϊ��ϵͳ�İ�ȫ�ԣ��뿪���������˳�ϵͳ</font> </p></td>
  </tr>
  <tr> 
    <td width="100%" bgcolor="#3CC49E">&nbsp;</td>
  </tr>
</table>
</html>