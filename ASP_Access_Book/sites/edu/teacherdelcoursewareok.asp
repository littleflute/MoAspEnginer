<!--#include file="conn.asp"-->
<!--#include file="isteacher.asp"-->
<%
id = trim(request("id"))
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('�벻Ҫ����');top.window.location.href='teachermain.asp';</script>"
 response.end
end if

sql = "select * from main where mainid="&id
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,1,1
if rs.bof and rs.eof then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('�벻Ҫ����');top.window.location.href='teachermain.asp';</script>"
 response.end
elseif rs("idofteacher") <> int(session("teacherid")) then
 rs.close
 set rs = nothing
 conn.close
 set conn = nothing
 response.write "<script>alert('������ϲ����㷢���ģ������ʲô��');top.window.location.href='teachermain.asp';</script>"
 response.end
else
 fileurl = rs("fileurl")
 rs.close
 set rs=nothing
'����������ϴ������أ���ɾ������
 if left(fileurl,6) = "files/" and mid(fileurl,7,len(session("teacherid"))+2) = session("teacherid")&"at" then
  dim filepaths,objFSO
  on error resume next
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  filepaths=Server.MapPath(""&fileurl&"")
  if objFSO.fileExists(filepaths) then
   objFSO.DeleteFile(filepaths)
  end if
  set objFSO = nothing
 end if
 conn.execute "delete from main where mainid="&id
 conn.close
 set conn = nothing
 response.write "<script>alert('ɾ���ɹ�');window.location.href='list.asp?teacherid="&session("teacherid")&"';</script>"
end if
%>