<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!-- #include file="an.htm" -->
<table width="760" border="0" align="center">
  <tr> 
    <td colspan="2" bgcolor="#FFFF99"><div align="center">�û�ע��</div></td>
  </tr>
  <tr> 
    <td colspan="2">
	<form ACTION="addsave.asp" METHOD="POST" name="form" id="form">
        <table width="95%" border="1" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
          <tr> 
            <td width="24%" bgcolor="#CCCCCC"> <div align="right">�û�����</div></td>
            <td width="76%"><input name="id" type="text" id="id"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">�� �룺</div></td>
            <td> <input name="psw" type="password" id="psw"></td>
          </tr>
          <tr> 
            <td height="18" bgcolor="#CCCCCC"> <div align="right">�� ��</div></td>
            <td><input name="sex" type="radio" value="boy" checked>
              ˧�� &nbsp; <input type="radio" name="sex" value="girl">
              ��Ů</td>
          </tr>
          <tr> 
            <td height="19" bgcolor="#CCCCCC"> <div align="right">Q Q��</div></td>
            <td><input name="qq" type="text" id="qq"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">E-mail��</div></td>
            <td><input name="email" type="text" id="email"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"><div align="right">������ҳ��</div></td>
            <td><input name="homepage" type="text" id="homepage"> </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">�Ƿ�Э��Ա��</div></td>
            <td> <input name="huiyuan" type="radio" value="yes" checked>
              �� &nbsp; <input type="radio" name="huiyuan" value="no">
              ��</td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#CCCCCC">���½�����Э��ͨ��¼�������������д</td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">������</div></td>
            <td bgcolor="#CCCCCC"><input name="name" type="text" id="name"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">���ţ�</div></td>
            <td> <select name="part" id="part">
                <option value="�᳤">�᳤</option>
                <option value="���᳤">���᳤</option>
                <option value="���鴦">���鴦</option>
                <option value="��Ŀ��">��Ŀ��</option>
                <option value="���ز�">���ز�</option>
                <option value="������">������</option>
                <option value="OEC���">OEC���</option>
                <option value="��Ա���ֲ�">��Ա���ֲ�</option>
              </select></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">ְ��</div></td>
            <td> <select name="class" id="class">
                <option value="�᳤">�᳤</option>
                <option value="���᳤">���᳤</option>
                <option value="����">����</option>
                <option value="������">������</option>
                <option value="��Ա">��Ա</option>
              </select> </td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">Ժϵ&amp;�꼶��</div></td>
            <td> <input name="college" type="text" id="college"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">�� ����</div></td>
            <td> <input name="tel" type="text" id="tel"></td>
          </tr>
          <tr> 
            <td bgcolor="#CCCCCC"> <div align="right">סַ��</div></td>
            <td> <input name="addr" type="text" id="addr"></td>
          </tr>
        </table>
        <p align="center"> 
          <input type="submit" name="Submit" value="  �����  ">
          <input type="reset" name="Submit2" value="  �� ��  ">
        </p>
        <input type="hidden" name="MM_insert" value="form1">
      </form></td>
  </tr>
  <tr> 
    <td width="242">&nbsp;</td>
    <td width="508">&nbsp;</td>
  </tr>
</table>
</body>
</html>
