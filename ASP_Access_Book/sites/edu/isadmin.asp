<%
if session("admin") <> "admin" then
 response.write "<script>alert('您不是管理员，或您尚未登陆');top.window.location.href='adminlogin.asp';</script>"
 response.end
end if
%>