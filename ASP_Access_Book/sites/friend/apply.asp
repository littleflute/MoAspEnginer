<%
Option Explicit
On Error Resume Next
%>
<!--#include file="conn.asp"-->
<%
dim JMail,body
dim rs_apply,rs_lar,rs_email
dim user_id,netname,sql,Email,myEmail,myname

'�Ѷ�Session�����Ƿ�ʱ
if isempty(session("user_id")) then
   response.redirect "timeout.htm"
end if

if session("user_id")=1 then
	response.redirect "notreg.htm"
	response.end
end if

user_id=request("user_id")
netname=request("netname")

'�����ݿ⣬���ʼ�

set rs_email=server.createobject("adodb.recordset")
'��ѯ��ǰ�Ự�ߵ�����
sql="select * from larchives where user_id=" & session("user_id")
rs_email.open sql,conn,3,2
if rs_email.eof and rs_email.bof then
	response.redirect "notreg.htm"
	response.end
end if
myEmail = rs_email("email")
myname = rs_email("name")

'----------------------------ȡ������
rs_email.close
sql="select * from larchives where user_id=" & user_id
rs_email.open sql,conn,3,2
if rs_email.eof and rs_email.bof then
	response.redirect "notreg.htm"
	response.end
end if

Email = rs_email("email")
body = "����:"&vbcrlf
body = body & vbTab & "��·���ｻ�����ĵġ�" & myname & "����������������������"&vbcrlf
body = body&"�뵽" & "" & "�鿴"&vbcrlf

'-------------------------------------------------ȡ������
rs_email.close
sql="select * from user_reg where user_id=" & user_id
rs_email.open sql,conn,3,2
if rs_email.eof and rs_email.bof then
	response.redirect "notreg.htm"
	response.end
end if
body = body&"�����û�����"& rs_email("user_name")&vbcrlf
body = body&"	�������룺"& rs_email("password")&vbcrlf
'body = body&"	Mailto:" & myemail
body = body&"-----------------------------------------------------------------------------------------"&vbcrlf
body=body&"��Ϊ�������ͨ����·���ｻ������ϵͳ�Զ���������,������Ҳ���ղ�����Ļ��ŵģ������벻Ҫֱ�ӻظ�����š�лл��"


Set JMail = Server.CreateObject("JMail.SMTPMail") 
JMail.ServerAddress = "202.104.147.132"
JMail.Sender = "webmaster@love.net"
JMail.Subject = "��������"
JMail.AddRecipient Email
JMail.Body = body
JMail.charset= ""
JMail.Priority = 3
JMail.AddHeader "Originating-IP", "10.11.1.12"
JMail.Execute 


set rs_apply=server.createobject("adodb.recordset")
sql="select * from apply where (for_id=" & session("user_id") & " and user_id=" & user_id & ") or (user_id=" & session("user_id") & " and for_id=" & user_id & ")"
rs_apply.open sql,conn,3,2
if not(rs_apply.eof and rs_apply.bof) then
	response.redirect "read.asp?user_id=" & user_id
	response.end
end if
rs_apply.close
set rs_apply=nothing

set rs_lar   =server.createobject("adodb.recordset")
sql="select * from larchives where user_id=" & cstr(session("user_id"))
rs_lar.open sql,conn,3,2
if rs_lar.eof and rs_lar.bof then
	response.redirect "notreg.htm"
	response.end
end if

set rs_apply=server.createobject("adodb.recordset")
rs_apply.open "apply",conn,3,2
rs_apply.addnew
rs_apply("for_id")=user_id
rs_apply("user_id")=session("user_id")
rs_apply("user_name")=rs_lar("name")
rs_apply("user_netname")=rs_lar("netname")
rs_apply("user_sex")=rs_lar("sex")
rs_apply("user_age")=rs_lar("age")
rs_apply("user_home")=rs_lar("home")
rs_apply("apply_date")=date
rs_apply.update
rs_apply.close
rs_lar.close
set rs_lar=server.createobject("adodb.recordset")
sql="select renqi from larchives where user_id =" & user_id
rs_lar.open sql,conn,3,2
rs_lar("renqi")=rs_lar("renqi")+1
rs_lar.update
rs_lar.close

set rs_lar=nothing
set conn=nothing
set rs_email = nothing

response.redirect "read.asp?user_id=" & user_id
%>
