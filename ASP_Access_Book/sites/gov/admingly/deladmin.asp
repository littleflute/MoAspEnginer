<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<%
         dim id
		 id=request.QueryString("id")
		 if id="" or isnull(id) then
		 response.write"<script>alert('对不起,操作符丢失\n请返回重试');history.back();</script>"
              else
			  sql="delete from admingly where id in ("&id&")"
			  conn.execute(sql)
	 %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>删除成功</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">请返回上一页后进行查看</div></td>
  </tr>
  <tr> 
    <td align="center"><a href="listalladmin.asp">【返回】</a></td>
  </tr>
</table>
<% end if %>
</body>
</html>
