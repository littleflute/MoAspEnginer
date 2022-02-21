<!--#include file="conn.asp"-->
<!--#include file="isadmin.asp"-->
<%
id = trim(request("id"))
returnlist = trim(request("returnlist"))
if id = "" then
 conn.close
 set conn = nothing
 response.write "<script>alert('请不要捣乱');top.window.location.href='adminmain.asp';</script>"
else
 conn.execute "delete from teacher where teacherid="&id
'删除该教师上传的资料
 sql = "select * from main where idofteacher="&id
 set rs = server.createobject("adodb.recordset")
 rs.open sql,conn,1,1
 if not (rs.bof and rs.eof) then
  dim filepaths,objFSO
  on error resume next
  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  do while not rs.eof
   fileurl = rs("fileurl")
   if left(fileurl,6) = "files/" and mid(fileurl,7,len(id)+2) = id&"at" then
    filepaths=Server.MapPath(""&fileurl&"")
    if objFSO.fileExists(filepaths) then
     objFSO.DeleteFile(filepaths)
    end if
   end if
   rs.movenext
  loop
  set objFSO = nothing
 end if
 rs.close
 set rs = nothing
 conn.execute "delete from main where idofteacher="&id
 conn.close
 set conn = nothing
 if returnlist = "yes" then
  response.write "<script>alert('删除成功');window.location.href='unlock.asp';</script>"
 else
  response.write "<script>alert('删除成功');window.location.href='adminteacher.asp';</script>"
 end if
end if
%>