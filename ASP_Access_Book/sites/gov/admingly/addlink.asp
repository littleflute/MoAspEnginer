<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%      dim action,webaddr,fnames
         action=request.querystring("action")
        if action="add" then
	     if request.form("Sumit")<>"�� ��" then
		 response.write"<script>alert('��Ҫ����\n��������');history.back();</script>" 
                 end if
		 fnames=request.form("fnames")
		 webaddr=request.form("webaddr")
		 sql="select * from aboutlink where webaddr='"&webaddr&"' or zwname='"&fnames&"'"
		      set rs=conn.execute(sql)
			  if not rs.eof then
			 response.write"<script>alert('�˵�λ�Ѿ������\n�뷵��');history.back();</script>" 
		         end if
			 if webpage(webaddr) then
		 sql="insert into aboutlink(webaddr,zwname) values('"&webaddr&"','"&fnames&"')"
               conn.execute(sql)
			   %>			   
<table width="100%" border="0" align="center">
  <tr>
    <th>��ӳɹ�</th>
  </tr>
  <tr>
    <td>�����Ѿ����,�뷵����ҳ���в鿴.��ѡ����Ҫ���صĵص�.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_link.asp">�����ء�</a><a href="adminbody.asp">������ǰһҳ��</a> 
        </div></td>
  </tr>
</table>
<%
          else
			  response.write"<script>alert('�Բ���,��ַ���Ϸ�\n�뷵��������д');history.back();</script>"
			    end if
				end if
			    if action="" or isnull(action) then
			   %>
<form name="form1" method="post" action="addlink.asp?action=add">			   
  <table width="350" border="0" align="center">
    <tr> 
    <th colspan="2">�������</th>
  </tr>
  <tr> 
    <td>���ӵĵ�λ����:</td>
    <td>
        <input type="text" name="fnames">
      </td>
  </tr>
  <tr> 
    <td>���ӵ�λ����ַ:</td>
    <td><input type="text" name="webaddr"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><input type="submit" name="Sumit" value="�� ��"></td>
  </tr>
</table>
</form>
<% end if %>
</body>
</html>
