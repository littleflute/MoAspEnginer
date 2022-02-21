<html>
<head>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../xsmz.css" rel=stylesheet><title>无标题文档</title>
</head><body>
<%
if request("action") = "send" then
	frommail = Trim(Request.Form("frommail"))
	tomail = Trim(Request.Form("tomail"))
	mailsubject = Trim(Request.Form("mailsubject"))
	mailbody = Trim(Request.Form("mailbody"))
	 if frommail="" or tomail="" or mailsubject="" or mailbody="" then
	 response.write"<script>alert('对不起,请把资料填写详细');history.back();</script>"
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
		   response.write"<table width=300 border=1 align=center><tr><td align=center bgcolor='#B4E7F1'>邮件已发送</td>"&_
		   "</tr><tr><td>邮件已发送,您可以选择关闭此窗口</td></tr><tr><td align=center><a href='javascript:window.close();'>【关闭】</a></td></tr></table>"
	       Response.End
           end if
              %>
<form name="sendmail" action="?action=send" method="post" >
  <table border="0" cellspacing="1" cellpadding="4" width="400" align="center" bgcolor="#797161">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2" align="center">发 送 邮 件</td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td width="130" height="25" bgcolor="#f6f6f6">发信人地址：</td>
      <td width="251" bgcolor="#f6f6f6"> 
        <input type="text" name="frommail" >
        <br>
        （必需为有效的E-mail,否则发送不能成功）</td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="25" bgcolor="#f6f6f6">收信人地址：</td>
      <td height="25" bgcolor="#f6f6f6"> 
        <input name="tomail" type="text" value="xx@xx.com" readonly="true">
      </td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="25" bgcolor="#f6f6f6">信件标题：</td>
      <td height="25"> <input type="text" name="mailsubject" size="35"></td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td bgcolor="#f6f6f6">信件内容：</td>
      <td bgcolor="#f6f6f6"> <textarea name="mailbody" cols="30" rows="8"></textarea> 
        <br>
        *并不是所有邮箱都支持CDONTS的.<br> </td>
    </tr>
    <tr bgcolor="#f6f6f6"> 
      <td height="10" colspan="2" align="center"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
        <input type="submit" name="Submit" value="发送"> <input type="reset" name="Submit" value="取消"> 
      </td>
    </tr>
  </table>
      </form>
</body>
</html>
