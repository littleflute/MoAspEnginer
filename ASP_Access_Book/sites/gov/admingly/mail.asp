<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet><title>�ޱ����ĵ�</title>
</head><body>
<%
if request("action") = "send" then
	frommail = Trim(Request.Form("frommail"))
	tomail = Trim(Request.Form("tomail"))
	mailsubject = Trim(Request.Form("mailsubject"))
	mailbody = Trim(Request.Form("mailbody"))
	 if frommail="" or tomail="" or mailsubject="" or mailbody="" then
	 response.write"<script>alert('�Բ���,���������д��ϸ');history.back();</script>"
	    end if
	   if tomail<>"" then
		Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
		objCDOMail.From = frommail
		objCDOMail.To = tomail
		objCDOMail.Subject = mailsubject  
			objCDOMail.BodyFormat = 0
			objCDOMail.MailFormat = 0
		objCDOMail.Body = mailbody
		objCDOMail.Send
		Set objCDOMail = Nothing
	       end if  
		   response.write"<table width=300 border=1 align=center><tr><td align=center bgcolor='#B4E7F1'>�ʼ��ѷ���</td>"&_
		   "</tr><tr><td>�ʼ��ѷ���,������ѡ��رմ˴���</td></tr><tr><td align=center><a href='javascript:window.close();'>���رա�</a></td></tr></table>"
	       Response.End
           end if
              %>
<form name="sendmail" action="?action=send" method="post" >
  <table border="0" cellspacing="1" cellpadding="4" width="400" align="center" bgcolor="#797161">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2" align="center">�� �� �� ��</td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td width="130" height="25" bgcolor="#f6f6f6">�����˵�ַ��</td>
      <td width="251" bgcolor="#f6f6f6"> 
        <input type="text" name="frommail" >
        <br>
        ������Ϊ��Ч��E-mail,�����Ͳ��ܳɹ���</td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="25" bgcolor="#f6f6f6">�����˵�ַ��</td>
      <td height="25" bgcolor="#f6f6f6"> 
        <input name="tomail" type="text" value="xx@xx.com" readonly="true">
      </td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="25" bgcolor="#f6f6f6">�ż����⣺</td>
      <td height="25"> <input type="text" name="mailsubject" size="35"></td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td bgcolor="#f6f6f6">�ż����ݣ�</td>
      <td bgcolor="#f6f6f6"> <textarea name="mailbody" cols="30" rows="8"></textarea> 
        <br>
        *�������������䶼֧��CDONTS��.<br> </td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="10" colspan="2" align="center"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="submit" name="Submit" value="����"> <input type="reset" name="Submit" value="ȡ��"> 
      </td>
    </tr>
  </table>
      </form>
</body>
</html>
