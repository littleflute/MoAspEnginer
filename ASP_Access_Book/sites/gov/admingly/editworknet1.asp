<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action,id
	 id=request.QueryString("id")
	 action=request.querystring("action")
	 if action="" or isnull(action) then
	 sql="select * from worknet where id="&id
	    set rs=conn.execute(sql)
		if rs.eof then
		response.write"<script>alert('�Է���,���Ҫ�����Ķ���δ�ҵ�');history.back();</scrpt>"
              else
	   %>
<form name="frmCtoy" method="post" action="editworknet1.asp?action=edit&id=<%=id%>" >
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bordercolorlight="#000000">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>������������ļ�</strong></div></td></tr>
    <tr>
      <td width="22%">�����������:</td>
      <td height="25"> <input name="nettitle" type="text" size="30" value="<%=rs("nettitle")%>"></td>
    <tr>
	<tr>
      <td width="22%">����:</td>
      <td width="78%"> 
        <textarea name="netbody" cols="60" rows="20"><%=rs("netbody")%></textarea>
      </td>
    </tr>
	<tr><td>��������:</td><td><input name="workfile" type="text" size="30" value="<%=rs("filename")%>"></td></tr>
 <tr> 
      <td height="40">�ϴ��ĸ�����</td>
      <td height="25"> 
     <input type="text" name="filename" value="<%=rs("workfile")%>">
        ע��,����㲻��̫��Ϥ�Ļ�.�˴���ò�Ҫ�޸�.</td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="ȷ��"></td>
	</tr>
</table>
</form>
 <%  end if  
end if 
  if  action="edit" then
     dim nettitle,netbody,workfile,filename
	     nettitle=request.form("nettitle")
		 netbody=request.form("netbody")
		 workfile=request.form("filename")
		 filename=request.form("workfile")
		 if nettitle="" or netbody="" then
		 response.write"<script>alert('�뷵�ذ�������д��ϸ');history.back();</script>"
             else
			 netbody=server.HTMLEncode(netbody)
		     sql="update worknet set nettitle='"&nettitle&"',netbody='"&netbody&"',workfile='"&workfile&"',filename='"&filename&"',nettime='"&now()&"' where id="&id
		     conn.execute(sql)
		 %>
		 
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B4E7F1">�޸ĳɹ�</th>
  </tr>
  <tr>
    <td>����Է��ص���һҳ��鿴</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_worknet.asp">�����ء�</a></div></td>
  </tr>
</table>
<% end if
   end if %>
</body>
</html>