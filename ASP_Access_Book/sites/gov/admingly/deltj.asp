<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim id
	    id=request.querystring("id")
	 if id="" or isnull(id) then
	    response.write"<script>alert('对不起,操作数据丢失\n请返回重试');hisory.back();</script>"
      else
	     sql="delete from mztj where tjid in("&id&")" 
		 conn.execute(sql)   
%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">删除成功</th>
  </tr>
  <tr>
    <td>删除的数据将是永久丢失,操作时请谨慎处理</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_tj.asp">【返回】</a><a href="adminbody.asp">【返回继续操作】</a></div></td>
  </tr>
</table>
<% end if%>
</body>
</html>
