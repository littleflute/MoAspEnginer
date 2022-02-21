<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim i,delfile,id,filename,workfile
   i=request.querystring("i")
   id=request.querystring("id")
    application.lock
    application("filename"&i)=""
	 application.unlock
	delfile=application("workfile"&i)
	application.Lock
	application("workfile"&i)=""
       application.UnLock
   response.write"<script>alert('"&delbackdb(delfile)&"');</script>"
     application.lock
	 for i=0 to application("all")
	    if application("filename"&i)<>"" and application("workfile"&i)<>"" then
	        if filename="" then
			filename=application("filename"&i)
			workfile=application("workfile"&i)
			else
	 filename=filename&";"&application("filename"&i)
	 workfile=workfile&";"&application("workfile"&i)
	         end if
       application("workfile"&i)=""
	   application("filename"&i)=""
	   end if
	   next
	   application.UnLock
	   sql="update worknet set filename='"&filename&"',workfile='"&workfile&"' where id="&id
	     conn.execute(sql)
	   %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="25" bgcolor="#B4E7F1"> 
      <div align="center"><strong>删除动作已完成</strong></div></td>
  </tr>
  <tr>
    <td>数据删除后将是永久丢失,操作时请谨慎.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="editworknet.asp?id=<%=id%>">【返回】</a></div></td>
  </tr>
</table>
</body>
</html>
