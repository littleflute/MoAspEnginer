<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action
	 action=request.querystring("action")
	 id=request.querystring("id")
	 if action="" or isnull(action) then
	 sql="select id, title,body from affiche"
        set rs=conn.execute(sql)
		if not rs.eof then
%>
<form name="form1" method="post" action="admin_affiche.asp?action=edit&id=<%=rs("id")%>">
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> <div align="center"><strong>��ǰ����</strong></div></td>
    </tr>
    <tr> 
      <td width="69">����:</td>
      <td width="271"> <input name="title" type="text" id="title" value="<%=rs("title")%>"> </td>
    </tr>
    <tr> 
      <td>����:</td>
      <td><textarea name="body" cols="35" rows="8" id="body"><%=rs("body")%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="�� ��">
        </div></td>
    </tr>
  </table>
</form>
<%  end if
    else
	if action="edit" then
	dim title,body
	title=trim(request.Form("title"))
	body=server.HTMLEncode(trim(request.Form("body")))
	if body="" or title="" then
	response.write"<script>alert('���������д��ϸ\n��������');history.back();</script>"
         else
		   if id="" or isnull(id) then
		 sql="insert into affiche(title,body,aftime) values('"&title&"','"&body&"','"&now()&"')"
		      else
		 sql="update affiche set title='"&title&"', body='"&body&"', aftime='"&now()&"' where id="&id
		      end if
		   conn.execute(sql)
		   end if
        %>
		<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="25" bgcolor="#B4E7F1"> 
      <div align="center"><strong>�޸ĳɹ�</strong></div></td>
  </tr>
  <tr>
    <td>����Է�����һҳ����в鿴</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_affiche.asp">�����ء�</a></div></td>
  </tr>
</table>
<% end if
   end if %>
</body>
</html>
