<%@language=vbscript%>
<%
  if session("admin")="" then
  response.redirect "admin.asp"
  else
	if session("flag")>1 then
		response.write "<br><p align=center>��û�в�����Ȩ��</p>"
		response.end
	end if
  end if
  
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript>

function check()

{
  if(document.form1.username.value=="")
    {
      alert("�û���Ϊ��");
      return false;
    }
    
  if(document.form1.newpin.value=="")
    {
      alert("���벻��Ϊ��");
      return false;
    }
    
  if((document.form1.newpin.value)!=(document.form1.re_newpin.value))
    {
      alert("���벻ƥ��");
      return false;
    }

}

</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <p>&nbsp;</p>
  <form method="post" action="saveuser1.asp" name="form1" onsubmit="javascript:return check();">
    <table width="40%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr bgcolor="#CCCCCC"> 
        <td height="25"> 
          <div align="center"><font size="2">��������Ա</font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">�� �� ���� 
            <input type="text" name="username">
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">��ʼ���롡 
            <input type="password" name="newpin">
            </font> </div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center"><font size="2">ȷ�����롡 
            <input type="password" name="re_newpin">
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="30"> 
          <div align="center">Ȩ������ 
             </div>
        </td>
      </tr>
      <tr>
        <td height="30">
          <div align="center">
            <input type="radio" name="right_class" value="3">
            �û���
            <input type="radio" name="right_class" value="2" checked>
            ��ͨ����Ա��
            <input type="radio" name="right_class" value="1">
            �����û�</div>
        </td>
      </tr>
    </table>
    <p>
      <input type="submit" name="Submit" value="ȷ��">
    </p>
  </form>
  <p>&nbsp;</p>
  <p align=center></p> </div>
</body>
</html>
