<%@language=vbscript%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="guest_sub.asp" -->
<%if Request.Form("from") = "SaveReply" then 
   dim reply,gPassWord,Content,gbID
   gPassWord=Request.Form("PassWord")
   if gPassWord<>PassWord then
      Msg("<font color=#000000><li>��������,ֻ�й���Ա���ܻظ�����.</li></font>")
	  Response.End
   End if
    reply   =Request.form("reply")
	Content=Request.form("Content")
	Set Book=Server.CreateObject("ADODB.RecordSet")
	Book.open "Select * from guest where ID = " & request("ID"),conn,3,3
	if Book.eof and Book.bof then
             Msg("<li>���������Ҳ������Ϊ" & request("ID") & "�����ԣ���˲��ܽ��лظ�������</li>")
             Response.End 
	else
	         Book("�ظ�����") = reply
             Book("�ظ�����") = cstr(now())
             Book("�ظ�")     =Content
			 Book.update
	end if
    Redirect "guest_list.asp?PageNo=1","лл��Ļظ���2���ϵͳ���Զ����ء�"
   
 end if%>