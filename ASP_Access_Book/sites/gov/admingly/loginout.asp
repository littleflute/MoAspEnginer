<%
         session("username")=""
		 session("ispass")=""
		 response.Cookies("mzcoks")("unm")=""
		 response.Cookies("mzcoks")("upwd")=""
		 response.Write"<script>alert('session���Ƴ�\ncookies�����');window.location.href='adminlogin.asp';</script>"
%>
