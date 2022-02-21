<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim id
   id = Request.QueryString("id")
   if id="" then
   response.write"<script>alert('非法操作\n请返回重试');history.back();</script>"
      else
   sql="delete from stat where id in("&id&")"
     conn.execute(sql)
	 sql="delete from votes where fromid in("&id&")"
	 conn.execute(sql)
%>
<table width="100%" border="0">
  <tr>
    <th height="30">删除成功</th>
  </tr>
  <tr>
    <td>在此处删除,将是删除投票标题以前所有的投票选项!</td>
  </tr>
  <tr>
    <td><div align="right"><a href="javascript:history.go(-1)">【返回】</a></div></td>
  </tr>
</table>
 <% end if %>
</body>
</html>
