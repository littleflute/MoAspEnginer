<%
if session("admin") <> "admin" then
 response.write "<script>alert('�����ǹ���Ա��������δ��½');top.window.location.href='adminlogin.asp';</script>"
 response.end
end if
%>