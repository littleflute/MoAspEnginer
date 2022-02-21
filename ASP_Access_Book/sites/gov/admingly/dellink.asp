<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim id
   id=request.querystring("id")
   if id="" or isnull(id) then
   response.write"<script>alert('对不起,删除产生错误\n请返回重试');history.back();</script>"
          else
   sql="delete from aboutlink where id in("&id&")"
   conn.execute(sql)
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>删除成功</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">删除的文件将是永久删除,操作时请一定谨慎对待.请选择你返回要返回的页面.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="admin_link.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
