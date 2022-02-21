<!--#include file="ipjc.asp"-->
<%
     if  session("username")="" or session("ispass")<>"pass" then
		 response.Write"<script>alert('对不起，请先登陆');window.location.href='adminlogin.asp';</script>"
               end if
			   %>



