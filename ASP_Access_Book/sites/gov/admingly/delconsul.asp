<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id
	id=request.querystring("id")
	if id="" or isnull(id) then
	response.write"<script>alert('对不起,操作符丢失');history.back();</script>"
      else
	  sql="delete from contact where id in("&id&")"
	  conn.execute(sql)
	%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="25" bgcolor="#B1E1F3"> <div align="center"><strong>删除动作完成</strong></div></td>
  </tr>
  <tr> 
    <td>你可以返回上一页进行查看,为了尊重咨询者请尽量不要删除他的留言</td>
  </tr>
  <tr> 
    <td><div align="center"><a href="admin_consul.asp">【返回】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
