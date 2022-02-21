<!--#include file="mzfunc.asp"-->
<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="留言咨询" %>
<!--#include file="mztop.asp"-->
<%  dim guestname,from,homeaddr,tel,handset,guestcontent,errorcount
      dim times
	  errorcount=""
	  guestname=trim(request.Form("guestname"))
      from=trim(request.Form("from"))
	  homeaddr=trim(request.Form("homeaddr"))
	  tel=trim(request.Form("tel"))
	  handset=trim(request.Form("handset"))
	  guestcontent=trim(request.Form("guestcontent"))
	   if homeaddr="" or isnull(homeaddr) then
	     homeaddr="保密"
		 end if
	   if handset<>"" then
	  if not ishandset(handset) then
	  errorcount="手机号码您可以不用输入，但输入的话请输入正确的号码"
         end if
		 else
		 handset=000
		 end if
     if tel="" or isnull(tel) then
	 errorcount="请您务必留下电话号码以便于我们联系"
	   end if
	   if not isnumeric(tel) then
         errorcount="对不起,电话号码必需为数字"
		 end if
   if errorcount="" then
   	set rs = server.createobject("adodb.recordset")
   sql="select * from contact"
   rs.open sql,conn,1,3
   rs.addnew
   rs("guestname")=guestname
   rs("from")=from
   rs("homeaddr")=homeaddr
   rs("tel")=tel
   rs("handset")=handset
   rs("guestcontent")=guestcontent
   rs("times")=now()
   rs("reply")=""
    rs.update
    response.write"<table width=776 height=126 border=0 align=center><tr><td align=center bgcolor=#f1f1eb> 留言成功</td></tr>"&_
	"<tr><td align='center' bgcolor=#f1f1eb> <strong>您的申诉我们已受理,我们会尽快给您答复.</strong></td></tr>"&_
	"<tr><td align='center' bgcolor='#f1f1eb'> <a href='discontact.asp'>返回查看留言</a></td></tr></table>"
           else 
response.write"<script>alert('"&errorcount&"');history.back();</script>"
 end if 
 call mzfoot() 
%>