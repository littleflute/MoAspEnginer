<%
  	if session("admin")="" then
  		response.redirect "admin.asp"
  	end if
%>
<html>
<link rel="stylesheet" href="style.css">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#CCCCCC">
  <tr>
    <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
        <tr> 
          <td background="../c.jpg" bgcolor="#FFFFFF"><img src="../c.jpg" width="536" height="54"></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF"><div align="center"> 
              <table border="0" cellspacing="1" width="80%" align=center>
                <tr> 
                  <td width="100%"><br> <p align=center><b>��Ϣ��������ϵͳ</b></p>
                    <b>����Ա�ɽ��в���˵��</b>��<br> <br>
                    1��ͨ��Web������£���ѡ���������������û��������û� <br> <br>
                    2�����Ѿ���������޸Ļ�ɾ����������������ӽ��в����������û��������û�����ͨ����Ա <br> <br>
                    3������Ŀ��������޸�ɾ����������������ӽ��в����������û��������û� <br> <br>
                    4���Ե�������������޸�ɾ����������������ӽ��в����������û��������û� <br> <br>
                    5���û��������޸�ɾ����������������ӽ��в����������û��������û� <br> <br>
                    6��<font color=red>Ϊ��ϵͳ�İ�ȫ�ԣ��뿪���������˳�ϵͳ</font> <br> </td>
                </tr>
              </table>
            </div></td>
        </tr>
        <tr> 
          <td background="../b.jpg" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</html>