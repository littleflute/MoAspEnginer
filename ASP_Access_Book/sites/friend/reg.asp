<%
if session("user_id")<>1 then
   response.redirect "havereg.htm"
end if
%>
<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>��������</title>
<style>
<!--
-->
</style>
<script language="javascript">
<!--
function isok(theform)
{
if (theform.user_name.value.length<2 || theform.user_name.value.length>10)
  {
    alert("�û���������Ҫ2λ�����10λ��");
    theform.user_name.focus();
    return (false);
  }
if (theform.password.value.length<3 || theform.password.value.length>10)
  {
    alert("��������Ҫ3λ�����10λ�� ��");
    theform.password.focus();
    return (false);
  }
if (theform.password_two.value=="")
  {
    alert("���ܻ���ȷ������û���أ�");
    theform.password_two.focus();
    return (false);
  }

if (theform.password_two.value!=theform.password.value)
  {
    alert("����������ô��һ����");
    theform.password_two.focus();
    return (false);
  }

return (true);
}
-->
</script>
<link rel="stylesheet" href="ysb.css">
</head>

<body>
<form method="POST" action="regsubmit.asp" onsubmit="return isok(this)">
<table border="0" width="402" cellspacing="0" cellpadding="0" align="center">
    <tr align="center"> 
      <td width="400" height="25"><b>�������ע��</b></td>
  </tr>
  <tr>
      <td width="400" height="20"><font color="#FF0000">������±��ύ������ס�û��������룬�����Ժ��½ʱ��</font><font color="#008000"><a href="#">�����֤</a></font></td>
  </tr>
</table>
  <table border="1" width="400" bordercolor="#CC0000" cellspacing="0" cellpadding="0" height="66" align="center" bgcolor="#067B0F">
    <tr>
      <td width="100%" valign="top" height="43">
        <table border="0" width="100%" cellspacing="0" cellpadding="0" bgcolor="#FFEBE1">
          <tr>
            <td width="43%" height="30"> 
              <p align="center">�û�����:</td>
            <td width="57%" height="30"> 
              <input type="text" name="user_name" size="18">(2-10)</td>
          </tr>
          <tr>
            <td width="43%" height="30"> 
              <p align="center">��������:</td>
            <td width="57%" height="30">
<input type="password" name="password" size="18">(3-10)</td>
          </tr>
          <tr>
            <td width="43%" height="30"> 
              <p align="center">ȷ������:</td>
            <td width="57%" height="30">
<input type="password" name="password_two" size="18"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<p align="center">��<input type="submit" value="�ύ" name="B1"><input type="reset" value="ȫ����д" name="B2"></p>
</form>

</body>

</html>
