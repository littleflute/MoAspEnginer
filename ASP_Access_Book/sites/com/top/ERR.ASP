
<%
 adminpwd=session("loginpwd")
 adminame=session("loginname")
%>
<!-- #include file="connect.asp" -->
<%
rs.Open "Select * From admin where admin_name='" &adminame&"'", conn, 3,3%>
<%if not (adminpwd=rs("admin_pwd") and adminame=rs("admin_name")) then%>
<script language=vbscript>  
  		MsgBox "对不起，您还没有正确登录或是未操作的时间过长！"  
  		location.href = "../login.asp"  
	  </script>
<%end if%>
<%rs.close
%>