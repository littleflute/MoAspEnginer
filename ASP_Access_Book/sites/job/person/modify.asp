<% Response.Buffer=True %>
<!--#include file="../inc/person.inc"-->
<% uname=session("puid")
if request("del")="true" then
conn.Execute("delete * from person where uname='"&uname&"'")
conn.Execute("delete * from pmailbox where reid='"&uname&"'")
conn.Execute("delete * from pfavorite where uname='"&uname&"'")
conn.Execute("delete * from cfavorite where fuid='"&uname&"'")
response.write"<SCRIPT language=JavaScript>alert('�û�ע���ɹ���');"
response.write"this.location.href='../';</SCRIPT>"
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'and job<>'""'"
rs.open sql,conn,1,1 %>
<%  if rs.eof or rs.bof then
   response.write"<SCRIPT language=JavaScript>alert('����δ��¼��ְ���������ȵ�¼��ְ������');"
   response.write"javascript:history.go(-1)</SCRIPT>"
   end if %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Content-Language" content="zh-cn">
<link rel="stylesheet" href="../inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;��ҳ</title>
</head>
<body topmargin="0" leftmargin="0">
<!--#include file="../inc/top2.htm"-->
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="149">
    <tr>
      <td width="778" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        �� </td>
    </tr>
    <tr>
      <td width="139" valign="top" bgcolor="#53BEB0"> �� 
        <div align="center">
          <center>
          <table border="0" cellpadding="0" cellspacing="0" width="118" height="280">
            <tr>
              <td width="100%" height="163" background="../images/stat-bg.GIF" valign="top" align="center">
        <p align="center"><br>
        <a href="main.asp">��¼��ҳ</a><br>
        <br>
        <a href="register.asp">��¼��ְ����</a><br>
        <br>
        <a href="modify.asp">������ְ����</a><br>
        <br>
        <a href="../changepwd.asp?stype=person" target="_blank">�޸ĵ�¼����</a><br>
        <br><a href="search.asp">ȫ��ְλ�б�</a>
              </td>
            </tr>
            <tr>
              <td width="100%" height="117" background="../images/stat-bg.GIF" valign="top" align="center">
        <p align="center"><br>
        <br><a href="favorite.asp">�ҵ��ղؼ�</a><br>
        <br>
        <a href="mailbox.asp">�ҵ�����</a><br>
        <br>
        <a href="../exit.asp">�˳���¼</a>
              </td>
            </tr>
          </table>
          </center>
        </div>
        <p align="center">&nbsp; </td>
      <td width="38" height="413" valign="top"><img border="0" src="../images/selfk.GIF"></td>
      <td width="455" height="413" valign="top">
  </center>
    
      <div align="center">
        <table border="0" cellpadding="0" cellspacing="0" width="241" height="195">
          <tr>
            <td width="239" height="148">
      <p align="left"><br>
      <br>
      ����Ҫ������ְ��������һ���֣�<br>
      <a href="register.asp?modify=ture"><br>
      </a><img border="0" src="../images/push.gif"> <a href="register.asp?modify=ture">���˻�������</a></p>   
      <p align="left"><img border="0" src="../images/push.gif"> <a href="register2.asp?modify=ture">������Ҫ�س�����ع�������</a> </p>    
      <p align="left"><img border="0" src="../images/push.gif"> <a href="register3.asp?modify=ture">ϣ��������������ϵ��Ϣ<br>    
      <br>
      </a></p>
            </td>
          </tr>
          <tr>
      <td height="1" valign="top" bgcolor="#000000">
      </td>
          </tr>
          <tr>
            <td width="239" height="48">
      <p align="left"><img border="0" src="../images/push.gif"> <a href="#"onclick="javascript:if (confirm('�Ƿ�ȷ��Ҫע���ʺ�?')) href='modify.asp?del=true'; 
	    else return;">ע���ʺ�</a>
            </td>
          </tr>
        </table>
      </div>
    
      <p>��
    
  </td>
  <center>
      <td width="1" height="413" valign="top" bgcolor="#00006A"></td>
      <td width="133" height="413" valign="top" bgcolor="#F3F3F3">��</td>
    </tr>
    <tr>
      <td width="778" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="778" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="../inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="778" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>
</body>

</html>
