<%
if session("admin") <> "admin" then
 if session("teacherid") = "" then
  conn.close
  set conn = nothing
  response.write "<script>alert('����δ��¼�����ѳ�ʱ��û�ж����������µ�¼');top.window.location.href='login.asp';</script>"
  response.end
 else
  sql = "select * from teacher where teacherid="&session("teacherid")
  set rsteacher = server.createobject("adodb.recordset")
  rsteacher.open sql,conn,1,1
  teacher = rsteacher("teacher")
  fenlei1 = rsteacher("fenlei1")
  fenlei2 = rsteacher("fenlei2")
  rsteacher.close
  set rsteacher = nothing
 end if
end if
%>