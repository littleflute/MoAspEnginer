<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
   id = Request.QueryString("id")
   if id="" or isnull(id) then
   response.write"<script>alert('�������ִ���\n�뷵������');history.back();</script>"
   else
sql = "DELETE FROM allarti WHERE id In (" & id & ")"
   conn.execute(sql)
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>ɾ���ɹ�</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">ɾ�����ļ���������ɾ��,����ʱ��һ�������Դ�.��ѡ���㷵��Ҫ���ص�ҳ��.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="sort_new.asp">�����ء�</a><a href="adminbody.asp">����������������</a></div></td>
  </tr>
</table>
<% end if %>
<!--#include file="adminfoot.asp"-->
</body>
</html>
