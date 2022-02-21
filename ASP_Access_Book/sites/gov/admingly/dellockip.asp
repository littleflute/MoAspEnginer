<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
     <%    
		 dim id
	        id=request.querystring("id")
			if id="" or isnull(id) then
			response.write"<script>alert('发生错误,操作符丢失\n请返回重试');history.back();</script>"
			       else
			 sql="delete from killip where id in ("&id&")"
			      conn.execute(sql)
				  end if
			%>
			<table width="350" border="1" align="center" bordercolor="#000000">
  <tr>
    <td height="25" align="center" bgcolor="#B1E1F3"><strong>删除成功</strong></td>
  </tr>
  <tr>
    <td>锁定ip只是防止黑客入侵的一种手段.请你经常备份数据库,以免造成不必要的损失.</td>
  </tr>
  <tr>
    <td align="center"><a href="unlock.asp">【返回】</a></td>
  </tr>
</table>	 
</body>
</html>
