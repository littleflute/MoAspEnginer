<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim id
	    id=request.querystring("id")
	 if id="" or isnull(id) then
	    response.write"<script>alert('�Բ���,�������ݶ�ʧ\n�뷵������');hisory.back();</script>"
      else
	     sql="delete from mztj where tjid in("&id&")" 
		 conn.execute(sql)   
%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">ɾ���ɹ�</th>
  </tr>
  <tr>
    <td>ɾ�������ݽ������ö�ʧ,����ʱ���������</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_tj.asp">�����ء�</a><a href="adminbody.asp">�����ؼ���������</a></div></td>
  </tr>
</table>
<% end if%>
</body>
</html>
