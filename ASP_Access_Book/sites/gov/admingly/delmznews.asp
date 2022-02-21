<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%  
    dim id
	    id=request.querystring("id")
        if id="" or isnull(id) then
		   response.write"<script>alert('操作符丢失,\n请返回重试');history.back();</script>"
           else
		   sql="delete from news where id in ("&id&")"
		       conn.execute(sql)
         %>
		 <table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">删除成功</th>
  </tr>
  <tr>
    <td>数据已成功删除,你可以返回查看.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_news.asp">【返回】</a> </div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
