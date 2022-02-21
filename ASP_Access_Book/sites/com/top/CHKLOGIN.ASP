
<%
adminame=trim(request.form("username"))
adminpwd=trim(Request.Form("password"))
if adminame="" then%>
<script language=vbscript>  
  		MsgBox "错误：请输入用户名字！"  
  		location.href = "javascript:history.back()"  
	  </script> 
	  <%
  end if
if adminpwd="" then%><script language=vbscript>  
  		MsgBox "错误：请输入用户密码！"  
  		location.href = "javascript:history.back()"  
	  </script> <%
end if%>
<!-- #include file="connect.asp" -->
<%rs.Open "Select * From admin where name='" &adminame&"'", conn, 3,3%>
<%if rs.bof then %><script language=vbscript>  
  		MsgBox "错误：此用户名不存在！"  
  		location.href = "javascript:history.back()"  
	  </script><%
elseif adminpwd<>rs("password") then%><script language=vbscript>  
  		MsgBox "错误：您的密码不正确！"  
  		location.href = "javascript:history.back()"  
	  </script><%
 else
 session("loginpwd")=rs("password")
 session("loginname")=rs("name")
    response.redirect "add.asp"
end if
rs.close
set rs=nothing%>
