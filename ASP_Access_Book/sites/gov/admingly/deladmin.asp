<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<%
         dim id
		 id=request.QueryString("id")
		 if id="" or isnull(id) then
		 response.write"<script>alert('�Բ���,��������ʧ\n�뷵������');history.back();</script>"
              else
			  sql="delete from admingly where id in ("&id&")"
			  conn.execute(sql)
	 %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>ɾ���ɹ�</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">�뷵����һҳ����в鿴</div></td>
  </tr>
  <tr> 
    <td align="center"><a href="listalladmin.asp">�����ء�</a></td>
  </tr>
</table>
<% end if %>
</body>
</html>
