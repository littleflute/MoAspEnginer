<%@ Language=VBScript %>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="guest_sub.asp" -->
<%Function Adjust_dapi(adj_str)
  dim final_str,i
  adj_str=Trim(adj_str)
  final_str=""
  If Len(adj_str)>0 then
     for i=1 to Len(adj_str)
	 Select Case Mid(adj_str,i,1)
	   Case Chr(13):final_str=final_str & ""
	   Case Chr(10):final_str=final_str & ""
	   Case Else:final_str=final_str & Mid(adj_str,i,1)
	 End Select
	 Next
  End if
  Adjust_dapi=final_str
End Function%>

<%
dim ASPBook
dim StrSQL
if isempty(Request("form")) then
       Msg("<font color=red size=2><b><li>对不起，没有操作请求，因此不能使用。</li></b></font)")
       Response.End
else
    dim UserName,HomePage,Email,Subject,Ly,Face,bookdate,LzHf,booktype,sql,Ly_str
     UserName      =Request.Form("UserName")
     HomePage      =Trim(Request.Form("HomePage"))
     Face          =Request.Form("Face")
     LzHf          =Request.Form("LzHf")
     Email         =Request.Form("Email")
     Subject       =Request.Form("Subject")
     Ly            =Request.Form("Ly")
	 Ly_str=Adjust_dapi(Ly)
	 IPinfo = Request.servervariables("REMOTE_ADDR")
     Set bc = Server.CreateObject("MSWC.BrowserType")   
     bookdate =Cstr(now())
     booktype= bc.browser + " " + bc.version + " " + bc.platform
     if Email<>"" then
	   if instr(Email,".")<=1 then
	      Msg("<font color=#000000><li>你的E-Mail有误,因此不能提交,请核对.</li></font>")
	      Response.End
	   end if
	   if Len(Email)<6 then
	     Msg("<font color=#000000><li>你的E-Mail有误,因此不能提交,请核对好后再试</li></font>")
	     Response.End
	   end if
	 end if
	 if UserName<>"" then
	   if Len(UserName)>255 then
	   Msg("<font color=#000000><li>你大名太长了,因此不能提交,请缩减一些.</li></font>")
	   Response.End
	   end if
	 end if
	 if HomePage<>"http://"then
	     if instr(HomePage,".")<=1 then
		   Msg("<font color=#000000><li>你的主页地址有误,因此不能提交,请核对后再试.</li></font>")
		   Response.End
		 end if
		 if Len(HomePage)<15 then
		   Msg("<font color=#000000><li>你的主页地址有误,因此不能提交,请核对后再试.</li></font>")
		   Response.End
		 end if
	 end if
	 if isnumeric(Ly) then
	   Msg("<font color=#000000><li>留言怎么光填数字呢?我们可不是华罗庚呀!</li></font>")
	   Response.End
	 end if  
	 
	 sql="Insert Into guest (姓名,来自,肖像,主页,邮件,主题,留言,IP,系统,留言日期)  Values('"& UserName &"','"& LzHf &"','"& Face &"','"& HomePage &"', '"& Email &"','"& Subject &"','"& Ly_str &"','"& IPinfo &"','"& booktype &"','"& bookdate &"')"
	 Set rs=conn.Execute(sql)
     Redirect "guest_list.asp?PageNo=1","谢谢你的留言！系统2秒后将自动返回。"
end if
%>	