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
       Msg("<font color=red size=2><b><li>�Բ���û�в���������˲���ʹ�á�</li></b></font)")
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
	      Msg("<font color=#000000><li>���E-Mail����,��˲����ύ,��˶�.</li></font>")
	      Response.End
	   end if
	   if Len(Email)<6 then
	     Msg("<font color=#000000><li>���E-Mail����,��˲����ύ,��˶Ժú�����</li></font>")
	     Response.End
	   end if
	 end if
	 if UserName<>"" then
	   if Len(UserName)>255 then
	   Msg("<font color=#000000><li>�����̫����,��˲����ύ,������һЩ.</li></font>")
	   Response.End
	   end if
	 end if
	 if HomePage<>"http://"then
	     if instr(HomePage,".")<=1 then
		   Msg("<font color=#000000><li>�����ҳ��ַ����,��˲����ύ,��˶Ժ�����.</li></font>")
		   Response.End
		 end if
		 if Len(HomePage)<15 then
		   Msg("<font color=#000000><li>�����ҳ��ַ����,��˲����ύ,��˶Ժ�����.</li></font>")
		   Response.End
		 end if
	 end if
	 if isnumeric(Ly) then
	   Msg("<font color=#000000><li>������ô����������?���ǿɲ��ǻ��޸�ѽ!</li></font>")
	   Response.End
	 end if  
	 
	 sql="Insert Into guest (����,����,Ф��,��ҳ,�ʼ�,����,����,IP,ϵͳ,��������)  Values('"& UserName &"','"& LzHf &"','"& Face &"','"& HomePage &"', '"& Email &"','"& Subject &"','"& Ly_str &"','"& IPinfo &"','"& booktype &"','"& bookdate &"')"
	 Set rs=conn.Execute(sql)
     Redirect "guest_list.asp?PageNo=1","лл������ԣ�ϵͳ2����Զ����ء�"
end if
%>	