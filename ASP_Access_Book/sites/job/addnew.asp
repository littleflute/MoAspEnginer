<% Response.Buffer=True %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/register.css" type="text/css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����˲š�&gt;���û�ע��</title>
</head>

<body topmargin="0" leftmargin="0">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<tr>
    
  <td height="33">&nbsp; </td>
  </tr>
      

<div align="center">
  <table width="780" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td><img src="inc/a3.jpg" width="780" height="148"></td>
    </tr>
  </table>
  <table border="0" cellpadding="0" cellspacing="0" width="780">
    <center>
      <tr>
        <td height="2" valign="top" colspan="3" bgcolor="#53BEB0"></td>
      </tr>
    </center>
    <tr>
      <td width="157" height="16" bgcolor="#53BEB0" valign="bottom"><div align="center"><a href="default.asp">��ҳ</a></div></td>
      <center>
        <td width="492" height="16" bgcolor="#53BEB0" valign="bottom">&nbsp;
            <marquee onmouseout=this.start() onmouseover=this.stop() border="0"  align="middle" scrolldelay="150" behavior="scroll"  width="100%">
            ��վ������
          </marquee></td>
      </center>
      <td width="131" height="16" bgcolor="#53BEB0" valign="bottom"><p align="center">
          <script language="javascript" src="inc/date.js"></script>
        </td>
    </tr>
    <center>
      <tr>
        <td height="2" valign="top" colspan="3" bgcolor="#53BEB0"></td>
      </tr>
    </center>
  </table>
</div>
 
<!-- #include file="inc/dbconn.inc" -->
<!-- #include file="inc/enpasswd.inc" -->
<SCRIPT language=JavaScript src="inc/validate.js"></SCRIPT>
<SCRIPT language=JavaScript>
<!--
function checkform()
{
 if (checkstring("�û���", document.addnew.uname.value, false)) {
    document.addnew.uname.focus();
    return false;   
  }
  var pwd = document.addnew.pwd.value;
  if (addnew.pwd.value=="") {
		alert("�������¼���룡");
		addnew.pwd.focus();		
		    return (false);
  }
  if (pwd.length < 3) {
    alert("���벻��������λ��");
    return false;   
  }
  var pwd2 = document.addnew.pwd2.value;
  if (pwd != pwd2) {
    alert("�����������벻һ�£�");
    document.addnew.pwd.value="";
    document.addnew.pwd2.value="";
	return false;
  }
  if (checkemail("�����ʼ�", document.addnew.email.value, false)) {
    document.addnew.email.focus();
    return false;   
  }
  return true;
}
//-->
</SCRIPT>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="780" height="45">
    <tr>
      <td width="766" height="18" valign="top" colspan="5" bgcolor="#53BEB0"> 
        �� </td>
    </tr>
	<tr>
      <td width="144" height="162" valign="top" bgcolor="#53BEB0">
        <p align="center">��</p>
        <p>��
      </td>
      <td width="29" height="309" valign="top"><img border="0" src="images/selfk.GIF"></td>
      <td width="462" height="309">
      <FORM name=addnew onsubmit="return checkform();" action=addnew.asp method=post>
  </center>
        <div align="center">
        <table border="1" cellpadding="0" cellspacing="0" width="352" height="1" bordercolor="#000000" bordercolorlight="#000000" bordercolordark="#FFFFFF">
          <tr>
            <td width="350" height="13" background="images/t-bg1.gif">
              <p align="center">=== ���û�ע�� ===</td>            
          </tr>
  <center>
          <center>
          <tr>
            <td width="350" height="263" bgcolor="#fffff4">
              <p align="center">&nbsp;&nbsp;&nbsp; �û���:<input type="text" name="uname" size="15" maxLength=15> </p>           
              <p align="center">&nbsp; ��¼����:<input type="password" name="pwd" size="15" maxLength=15></p>            
              <p align="center">&nbsp; �ظ�����:<input type="password" name="pwd2" size="15" maxLength=15></p>          
              <p align="center">��ϵE-mail:<input type="text" name="email" size="15" maxLength=30></p>
              <p align="center"><input type="radio" value="person" checked name="usertype">�����û�                                                       
                <input type="radio" name="usertype" value="company">��λ�û�</p>
              <p align="center"><input type="submit" value="ע ��" name="B1" style="position: relative; height: 19"></td>
          </tr>
          </center>
  </center>
  <center>
          <center>
          </center>
  </center>
        </table>
        </div>
      </td>
      
  <center>
      
	  <td width="1" height="309" valign="top" bgcolor="#00006A"></td>
      
	  <td width="119" height="309" valign="top" bgcolor="#F3F3F3"></td>
    </tr>
    <tr>
      <td width="743" height="1" valign="top" colspan="5" bgcolor="#53BEB0"> 
        <p align="center"> </td>
    </tr>
    <tr>
      <td width="743" height="3" valign="top" colspan="5">
      <p align="center">
      </td>
    </tr>
    <tr>
      <td width="743" height="7" valign="top" colspan="5">
      <p align="center"><script language="javascript" src="inc/copyright.js"></script>
      </td>
    </tr>
    <tr>
      <td width="743" height="3" valign="top" colspan="5">
      </td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

<% zhmail="��ӭ��ע�������˲ŵ��˲��г�,�뾡�췢��������Ƹ��Ϣ����ְ�������������ע��!"
zhmail=zhmail&"��ȫע��֮��,���������ܵ���վ���Ƶ���ְ��Ƹ����,Ŀǰ��Ҫ�����и����ղؼк�վ�����书��!"
zhmail=zhmail&"���ף���ڱ�վ�ҵ����ʵĹ������е����ʵ��˲�!"
zhmail=zhmail&"<br>&nbsp;&nbsp;&nbsp;&nbsp;P.S.(����Ϊϵͳ����,�벻Ҫ�ظ�,лл!)"
usertype=request("usertype") 
if usertype="" then Response.End
uname=request("uname")
email=request("email")
pwd=mistake(request("pwd"))

if usertype="person" then
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where uname='"&uname&"'"
rs.open sql,conn,3,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�û����ظ���������ѡ��һ���û�����');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if
rs.close
sql="select * from person"
rs.open sql,conn,3,3
rs.addnew
rs("uname")=uname
rs("pwd")=pwd
rs("email")=email
rs.update
rs.close
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from pmailbox"
rs.open sql,conn,3,3
rs.addnew
rs("reid")=uname
rs("senduid")="sysop"
rs("title")="ף����ע��ɹ���"
rs("sendname")="ϵͳ����Ա"
rs("sdate")=now()
rs("mailtext")=zhmail
rs.update
rs.close
session("puid")=uname
response.write"<SCRIPT language=JavaScript>alert('���û�ע��ɹ������ڵ�¼��...�����Ժ�');"
response.write"this.location.href='person/main.asp';</SCRIPT>"

else

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from company where uname='"&uname&"'"
rs.open sql,conn,3,3
if not rs.eof then
response.write"<SCRIPT language=JavaScript>alert('�û����ظ���������ѡ��һ���û�����');"
response.write"javascript:history.go(-1)</SCRIPT>"
end if
rs.close
sql="select * from company"
rs.open sql,conn,3,3
rs.addnew
rs("uname")=uname
rs("pwd")=pwd
rs("email")=email
rs.update
rs.close
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from cmailbox"
rs.open sql,conn,3,3
rs.addnew
rs("reid")=uname
rs("senduid")="sysop"
rs("title")="ף����ע��ɹ���"
rs("sendname")="ϵͳ����Ա"
rs("sdate")=now()
rs("mailtext")=zhmail
rs.update
rs.close
session("cuid")=uname
response.write"<SCRIPT language=JavaScript>alert('���û�ע��ɹ������ڵ�¼��...�����Ժ�');"
response.write"this.location.href='company/main.asp';</SCRIPT>" 
end if %>
