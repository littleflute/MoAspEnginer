<!--#include file="mzfunc.asp"-->
<!--#include file="rscoon.asp"-->
<!--#include file="totalcount.asp"-->
<% toptitle="������ѯ" %>
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
	     homeaddr="����"
		 end if
	   if handset<>"" then
	  if not ishandset(handset) then
	  errorcount="�ֻ����������Բ������룬������Ļ���������ȷ�ĺ���"
         end if
		 else
		 handset=000
		 end if
     if tel="" or isnull(tel) then
	 errorcount="����������µ绰�����Ա���������ϵ"
	   end if
	   if not isnumeric(tel) then
         errorcount="�Բ���,�绰�������Ϊ����"
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
    response.write"<table width=776 height=126 border=0 align=center><tr><td align=center bgcolor=#f1f1eb> ���Գɹ�</td></tr>"&_
	"<tr><td align='center' bgcolor=#f1f1eb> <strong>������������������,���ǻᾡ�������.</strong></td></tr>"&_
	"<tr><td align='center' bgcolor='#f1f1eb'> <a href='discontact.asp'>���ز鿴����</a></td></tr></table>"
           else 
response.write"<script>alert('"&errorcount&"');history.back();</script>"
 end if 
 call mzfoot() 
%>