<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action,id,webaddr,fnames
	 action=request.querystring("action")
	 id=request.querystring("id")
	       if id="" or isnull(id) then
	 response.write"<script>alert('�Բ���\n����������δ֪����');history.back();</script>"
               end if
		   if action="" then
		   sql="select webaddr,zwname,id from aboutlink where id="&id
		   set rs=conn.execute(sql) %>
  <form name="form1" method="post" action="edit_link.asp?action=edit&id=<%=rs("id")%>">			   
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#B1E1F3"> 
      <th height="30" colspan="2">�༭����</th>
  </tr>
  <tr> 
    <td>���ӵĵ�λ����:</td>
    <td>
        <input name="fnames" type="text" value="<%=rs("zwname")%>" size="30">
      </td>
  </tr>
  <tr> 
    <td>���ӵ�λ����ַ:</td>
    <td><input name="webaddr" type="text" value="<%=rs("webaddr")%>" size="30"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><input type="submit" name="Sumit" value="�� ��"></td>
  </tr>
</table>
</form>
<% end if 
   if action="edit" then
   webaddr=trim(request.form("webaddr"))
   fnames=trim(request.form("fnames"))
   if webaddr="" or fnames="" then
   response.write"<script>alert('�Բ���,�뽫���������������ȫ');history.back();</script>"
        end if
		  if not webpage(webaddr) then
  response.write"<script>alert('�Բ���,��ַ����ȷ,�����');history.back();</script>"
               else
   sql="update aboutlink set webaddr='"&webaddr&"', zwname='"&fnames&"' where id="&id			   
			   conn.execute(sql)
			  %>
			  
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">�޸ĳɹ�</th>
  </tr>
  <tr>
    <td>�޸�ǰ�����ݽ������Զ�ʧ,���޸�ʱ���������</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_link.asp">�����ء�</a></div></td>
  </tr>
</table>

			 <%  end if 
			     end if %>
</body>
</html>
