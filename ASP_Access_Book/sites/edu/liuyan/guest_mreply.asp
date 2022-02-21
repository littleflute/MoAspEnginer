<%@language=vbscript%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="guest_sub.asp" -->
<%if Request.Form("from") = "SaveReply" then 
   dim reply,gPassWord,Content,gbID
   gPassWord=Request.Form("PassWord")
   if gPassWord<>PassWord then
      Msg("<font color=#000000><li>密码有误,只有管理员才能回复留言.</li></font>")
	  Response.End
   End if
    reply   =Request.form("reply")
	Content=Request.form("Content")
	Set Book=Server.CreateObject("ADODB.RecordSet")
	Book.open "Select * from guest where ID = " & request("ID"),conn,3,3
	if Book.eof and Book.bof then
             Msg("<li>操作错误，找不到序号为" & request("ID") & "的留言，因此不能进行回复操作！</li>")
             Response.End 
	else
	         Book("回复主题") = reply
             Book("回复日期") = cstr(now())
             Book("回复")     =Content
			 Book.update
	end if
    Redirect "guest_list.asp?PageNo=1","谢谢你的回复！2秒后系统将自动返回。"
   
 end if%>