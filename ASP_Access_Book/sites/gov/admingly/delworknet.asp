<!--#include file="../mzfunc.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim id,delfile,i,j,allfile
	 	 id=request.querystring("id")
	 if  id="" or isnull(id) then
	      response.write"<script>alert('�Բ���,��������ʧ,\n��ҳ�治֧��ˢ��');history.back();</script>"
	     else
	 sql="select workfile from worknet where id in ("&id&")"
	   set rs=conn.execute(sql)
	   end if
	   delfile=rs.getrows
       for i=0 to ubound(delfile,2)
	     if delfile(0,i)<>"" or isnull(delfile(0,i))<>true then
              allfile=split(delfile(0,i),";")
			  for j=0 to ubound(allfile)
			    response.write"<script>alert('"&delbackdb(allfile(j))&"');</script>"
                       next
					   end if
					   next
			sql="delete from worknet where id in ("&id&")"
			    conn.execute(sql)
			  %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="25" bgcolor="#B1E1F3">ɾ���������</th>
  </tr>
  <tr>
    <td>ɾ�������ݽ��������Զ�ʧ,ɾ��ʱ���������</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_worknet.asp">�����ء�</a></div></td>
  </tr>
</table>
</body>
</html>
