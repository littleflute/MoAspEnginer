<!--#include file="conn.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="an.htm" -->
<table width="760" border="0" align="center" cellpadding="0" cellspacing="0" style="border: 1px #83BB37 solid; background-color: #e3f1d1;">
  <tr> 
    <td width="475" height="24"><a href="index.asp">��̳��̳</a>���û�ע��</td>
    <td width="283">|<a href="uerlist.asp">�û��б�</a> | <a href="adduer.asp">�û�ע��</a>| 
      <a href="find.asp">���һ�Ա</a>| <a href="uerlist.asp?order=2">��������</a>|</td>
  </tr>
</table>
<table width="760" border="0" align="center">
  <tr > 
    <td height="22" colspan="2" background="pic/h3.gif"><div align="center"><font color="#FFFFFF"><strong>�û�ע��</strong></font></div></td>
  </tr>
  <tr> 
    <td colspan="2">
	<form ACTION="addsave.asp" METHOD="POST" name="form" id="form">
        <table width="760" border="1" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
          <tr> 
            <td colspan="2" bgcolor="#e3f1d1">����˵������ע����Ϣͬʱ�����Э��ͨѶ¼����������ʵ��д</td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">�û���*��</div></td>
            <td width="465"><input name="id" type="text" id="id"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <p align="right">�� ��*��</td>
            <td width="465"> <input name="psw" type="password" id="psw"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">��������һ�����룺</div></td>
            <td width="465"> <input name="pswc" type="password" id="psw0" size="20"></td>
          </tr>
          <tr> 
            <td width="238" height="18" bgcolor="#e3f1d1"> <div align="right">�� 
                ��*��</div></td>
            <td width="465"><input name="sex" type="radio" value="boy" checked>
              ˧�� &nbsp; <input type="radio" name="sex" value="girl">
              ��Ů</td>
          </tr>
          <tr> 
            <td width="238" height="46" bgcolor="#e3f1d1"> <div align="right">���Ի�ͷ��<br>
              </div></td>
            <td width="465"><select                                  
      class=t2                                 
      onChange="document.images['idface'].src=options[selectedIndex].value;"                                 
      size=1 name=face>
                <option value=images/01.gif>�û�ͷ��-01 
                <option                
        value=images/02.gif>�û�ͷ��-02 
                <option value=images/03.gif>�û�ͷ��-03 
                <option                
        value=images/04.gif>�û�ͷ��-04 
                <option value=images/05.gif>�û�ͷ��-05 
                <option                
        value=images/06.gif>�û�ͷ��-06 
                <option value=images/07.gif>�û�ͷ��-07 
                <option                
        value=images/08.gif>�û�ͷ��-08 
                <option value=images/09.gif>�û�ͷ��-09 
                <option                
        value=images/10.gif>�û�ͷ��-10 
                <option value=images/11.gif>�û�ͷ��-11 
                <option                
        value=images/12.gif>�û�ͷ��-12 
                <option value=images/13.gif>�û�ͷ��-13 
                <option                
        value=images/14.gif>�û�ͷ��-14 
                <option value=images/15.gif>�û�ͷ��-15 
                <option                
        value=images/16.gif>�û�ͷ��-16 
                <option value=images/17.gif>�û�ͷ��-17 
                <option                
        value=images/18.gif>�û�ͷ��-18 
                <option value=images/19.gif>�û�ͷ��-19 
                <option                
        value=images/20.gif selected>�û�ͷ��-20</option>
              </select> <img src="images/20.gif" alt=����������� align="middle" class=t2 id=idface><font class=j1>[<a 
      href="tx.htm" target=_blank>����ͷ��</a>]</font></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">Q Q��</div></td>
            <td width="465"><input name="qq" type="text" id="qq2"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <p align="right">E-mail*��</td>
            <td width="465"><input name="email" type="text" id="email2"> </td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"><div align="right">������ҳ��</div></td>
            <td width="465"><input name="homepage" type="text" id="homepage2" value="http://"> 
            </td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">*�Ƿ�Э��Ա��</div></td>
            <td width="465"> <input name="huiyuan" type="radio" value="yes">
              �� &nbsp; <input type="radio" name="huiyuan" value="no">
              ��</td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#e3f1d1">�������½�����Э��ͨ��¼�������������д</td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">����*��</div></td>
            <td width="465"><input name="name" type="text" id="name"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">���ţ�</div></td>
            <td width="465"> <select name="part" id="part">
                <option value="�ǻ�Ա">�ǻ�Ա</option>
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
            <td width="238" bgcolor="#e3f1d1"> <div align="right">ְ��</div></td>
            <td width="465"> <select name="class" id="class">
                <option value="�ǻ�Ա">�ǻ�Ա</option>
                <option value="����">����</option>
                <option value="������">������</option>
                <option value="��Ա">��Ա</option>
                <option value="�᳤">�᳤</option>
                <option value="���᳤">���᳤</option>
              </select> </td>
          </tr>
          <tr> 
            <td bgcolor="#e3f1d1"><div align="right">Ժϵ&amp;�꼶��</div></td>
            <td><input name="college" type="text" id="college"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">�� ����</div></td>
            <td width="465"> <input name="tel" type="text" id="tel"></td>
          </tr>
          <tr> 
            <td width="238" bgcolor="#e3f1d1"> <div align="right">סַ��</div></td>
            <td width="465"> <input name="addr" type="text" id="addr"></td>
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
<!--#include file="foot.asp"-->