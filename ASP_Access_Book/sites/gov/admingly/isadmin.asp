<!--#include file="ipjc.asp"-->
<%
     if  session("username")="" or session("ispass")<>"pass" then
		 response.Write"<script>alert('�Բ������ȵ�½');window.location.href='adminlogin.asp';</script>"
               end if
			   %>



