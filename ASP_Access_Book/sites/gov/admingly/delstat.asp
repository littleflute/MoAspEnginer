<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim id
   id = Request.QueryString("id")
   if id="" then
   response.write"<script>alert('�Ƿ�����\n�뷵������');history.back();</script>"
      else
   sql="delete from stat where id in("&id&")"
     conn.execute(sql)
	 sql="delete from votes where fromid in("&id&")"
	 conn.execute(sql)
%>
<table width="100%" border="0">
  <tr>
    <th height="30">ɾ���ɹ�</th>
  </tr>
  <tr>
    <td>�ڴ˴�ɾ��,����ɾ��ͶƱ������ǰ���е�ͶƱѡ��!</td>
  </tr>
  <tr>
    <td><div align="right"><a href="javascript:history.go(-1)">�����ء�</a></div></td>
  </tr>
</table>
 <% end if %>
</body>
</html>
