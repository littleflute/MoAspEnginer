<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id
	id=request.querystring("id")
	if id="" or isnull(id) then
	response.write"<script>alert('�Բ���,��������ʧ');history.back();</script>"
      else
	  sql="delete from contact where id in("&id&")"
	  conn.execute(sql)
	%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="25" bgcolor="#B1E1F3"> <div align="center"><strong>ɾ���������</strong></div></td>
  </tr>
  <tr> 
    <td>����Է�����һҳ���в鿴,Ϊ��������ѯ���뾡����Ҫɾ����������</td>
  </tr>
  <tr> 
    <td><div align="center"><a href="admin_consul.asp">�����ء�</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
