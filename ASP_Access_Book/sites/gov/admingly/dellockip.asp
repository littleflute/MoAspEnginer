<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
     <%    
		 dim id
	        id=request.querystring("id")
			if id="" or isnull(id) then
			response.write"<script>alert('��������,��������ʧ\n�뷵������');history.back();</script>"
			       else
			 sql="delete from killip where id in ("&id&")"
			      conn.execute(sql)
				  end if
			%>
			<table width="350" border="1" align="center" bordercolor="#000000">
  <tr>
    <td height="25" align="center" bgcolor="#B1E1F3"><strong>ɾ���ɹ�</strong></td>
  </tr>
  <tr>
    <td>����ipֻ�Ƿ�ֹ�ڿ����ֵ�һ���ֶ�.���㾭���������ݿ�,������ɲ���Ҫ����ʧ.</td>
  </tr>
  <tr>
    <td align="center"><a href="unlock.asp">�����ء�</a></td>
  </tr>
</table>	 
</body>
</html>
