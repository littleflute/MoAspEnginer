<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim filename,i
  id = Request.QueryString("id")
  sql="select movies from video where id in("&id&")"
    set rs=conn.execute(sql)
	      filename=rs.getrows
  sql = "DELETE FROM video WHERE id In (" & id & ")"
  Conn.Execute(sql)
   for i=0 to ubound(filename,2)
   response.write"<script>alert('"&delbackdb(filename(0,i))&"')</script>"
   next
%>
<script language="javascript">
  alert("���ݿ��ļ��ѳɹ�ɾ����");
</script>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B1E1F3"> 
      <div align="center"><strong>ɾ���ɹ�</strong></div></td>
  </tr>
  <tr>
    <td><div align="center">ɾ�����ļ���������ɾ��,����ʱ��һ�������Դ�.��ѡ���㷵��Ҫ���ص�ҳ��.</div></td>
  </tr>
  <tr>
    <td height="20">
<div align="center"><a href="video.asp">�����ء�</a><a href="adminbody.asp">����������������</a></div></td>
  </tr>
</table>

</body>
</html>
