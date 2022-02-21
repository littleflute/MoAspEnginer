<%
         session("username")=""
		 session("ispass")=""
		 response.Cookies("mzcoks")("unm")=""
		 response.Cookies("mzcoks")("upwd")=""
		 response.Write"<script>alert('sessionÒÑÒÆ³ı\ncookiesÒÑÇå¿Õ');window.location.href='adminlogin.asp';</script>"
%>
