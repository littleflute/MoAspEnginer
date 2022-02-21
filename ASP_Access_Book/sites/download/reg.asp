<%
'---------------------����û�������-------------------------------
function Checkin(s) 
s=trim(s) 
s=replace(s," ","&amp;nbsp;") 
s=replace(s,"'","&amp;#39;") 
s=replace(s,"""","&amp;quot;") 
s=replace(s,"&lt;","&amp;lt;") 
s=replace(s,"&gt;","&amp;gt;") 
Checkin=s 
end function 

'---------------------����û�Email-------------------------------
function IsValidEmail(email)
IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if

end function

'---------------------�������-------------------------------
sub error()
%>
<!--#include file="dbpath.asp"-->


<html>
<head><title>������Ϣ��ʾ==<%=WebTitle%></title></head>
<body bgcolor="<%=bgcolor%>">
<div align="center">
  <table cellpadding=0 cellspacing=0 border=0 width=99% align=center>
    <tr>
      <td>
        <table cellpadding=3 cellspacing=1 border=0 width=100% bgcolor="<%=MainCColor%>">
          <tr align="center"> 
            <td width="100%">������Ϣ��</td>
          </tr>
          <tr> 
            <td width="100%"><b>��������Ŀ���ԭ��<br><%=errmsg%></td>
          </tr>
          <tr align="center">
            <td width="100%"><a target="_blank" href="index.asp"> << ��Ա��¼</a> ����<a href="Reg.asp">ע���Ա</a><input type="button" name="close" value="�رմ���" onclick="javascript:window.close()"></td>
          </tr>  
        </table>
      </td>
    </tr>
  </table>
</div>
</html>
<%
end sub

sub sendemail
	topic=FromUser&"��"&WebName&"�������һ�׸�"
	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
	mailbody=mailbody &"A:visited {	text-decoration: none;}"
	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TR><TD>"
	mailbody=mailbody &"��ã���������<a href="&WebUrl&" target=_blank>"&WebName&"</a>��ף����<br>"
	mailbody=mailbody &FromUser&"�ڱ�վ�������һ�׸�������<a href="&WebUrl&" target=_blank>����</a>��<br>"
	mailbody=mailbody &"</TD></TR></TABLE>"
	on error resume next
Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objcdo.MailFormat = 0
objcdo.BodyFormat = 1
objCDO.To         = ForUserEmail 
objCDO.Importance = 1
objCDO.From       = FromUserEmail
objCDO.Subject    = topic
objCDO.Body       = mailbody 
objCDO.Send
Set objCDO = Nothing
if err then 
	err.clear
else
	Response.Write "<center><b> �ż��Ѿ�������</b>"
end if


'MailFormat �ʼ��ĸ�ʽ��0��Html 1�����ı���
'BodyFormat ���ӵĸ�ʽ��1�����������Զ���Ϊ�ɵ����0����mailformat=0ʱ���Ӳ��䣬�����Ϊ�ɵ����
'To         �ʼ����շ��ĵ��������ַ
'Importance �ʼ�����Ҫ�ԣ�0���� 1���� 2���ߣ�
'From       �ʼ����ͷ��ĵ��������ַ
'Subject    �ʼ�������
'Body       �ʼ�������
'Send       ��ɷ����ʼ��Ķ���
'	topic=FromUser&"��"&WebName&"�������һ�׸�"
'	mailbody=mailbody &"<style>A:visited {	TEXT-DECORATION: none	}"
'	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none	}"
'	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline	}"
'	mailbody=mailbody &"A:link 	  {	text-decoration: none;}"
'	mailbody=mailbody &"A:visited {	text-decoration: none;}"
'	mailbody=mailbody &"A:active  {	TEXT-DECORATION: none;}"
'	mailbody=mailbody &"A:hover   {	TEXT-DECORATION: underline overline}"
'	mailbody=mailbody &"BODY   {	FONT-FAMILY: ����; FONT-SIZE: 9pt;}"
'	mailbody=mailbody &"TD	   {	FONT-FAMILY: ����; FONT-SIZE: 9pt	}</style>"
'	mailbody=mailbody &"<TABLE border=0 width='95%' align=center><TR><TD>"
'	mailbody=mailbody &"��ã���������<a href="&WebUrl&" target=_blank>"&WebName&"</a>��ף����<br>"
'	mailbody=mailbody &FromUser&"�ڱ�վ�������һ�׸�������<a href="&WebUrl&" target=_blank>����</a>��<br>"
'	mailbody=mailbody &"</TD></TR></TABLE>"
'	on error resume next
'    if EmailFlag=1 then 
'    	email = ForUserEmail
'		Dim JMail,SendMail
'		Set JMail=Server.CreateObject("JMail.SMTPMail")
'		JMail.Logging=True
'		JMail.Charset="gb2312"
'		JMail.ContentType = "text/html"
'		JMail.ServerAddress=SMTPServer
'		JMail.Sender=SystemEmail
'		JMail.Subject=topic
'		JMail.Body=mailbody
'		JMail.AddRecipient=email
'		JMail.Priority=3
'		JMail.Execute 
'		Set JMail=nothing 
'		if err then 
'			Response.Write "���������û�а�װJmail���������û�������ַ������ԭ�����޷�����"
'			err.clear
'		else
'			SendMail="OK"
'		    Response.Write "<center><b> �ż��Ѿ�������</b>"
'		end if
'	elseif EmailFlag=2 then
'    	email = useremail
'		dim  objCDOMail
'		Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
'		objCDOMail.From =SystemEmail
'		objCDOMail.To =email
'		objCDOMail.Subject =topic
'		objCDOMail.Body =mailbody
'		objCDOMail.Send
'		Set objCDOMail = Nothing
'		if err then 
'			Response.Write "���������û�а�װcdonts���������û�������ַ������ԭ�����޷�����"
'			err.clear
'		else
'		    Response.Write "<center><b> �ż��Ѿ�������</b>"
'		end if
'    elseif EmailFlag=3 then
'    	email = useremail
'		Set mailer=Server.CreateObject("ASPMAIL.ASPMailCtrl.1")  
'		recipient=email
'		sender=SystemEmail
'		subject=topic
'		message=mailbody
'		mailserver=SMTPServer
'		result=mailer.SendMail(mailserver, recipient, sender, subject, message)
'		if err then 
'			Response.Write "���������û�а�װaspmail���������û�������ַ������ԭ�����޷�����"
'			err.clear
'		else
'		    Response.Write "<center><b> �ż��Ѿ�������</b>"
'		end if
'    end if
end sub
%>