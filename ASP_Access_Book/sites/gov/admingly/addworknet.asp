<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action
	 action=request.querystring("action")
	 if action="" or isnull(action) then
%>
<form name="frmCtoy" method="post" action="addworknet.asp?action=save" >
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>������������ļ�</strong></div></td></tr>
    <tr>
      <td width="22%">�����������:</td>
      <td height="25"> <input name="nettitle" type="text" size="30"></td>
    <tr>
	<tr>
      <td width="22%">����:</td>
      <td width="78%"> 
        <textarea name="netbody" cols="60" rows="20"></textarea>
      </td>
    </tr>
	<tr><td>��������:</td><td><input name="workfile" type="text" size="30"></td></tr>
 <tr> 
      <td height="40">�ϴ��ļ���</td>
      <td height="25"> 
        <iframe frameborder="0" width="400" height=24 scrolling="no" src="upload1.asp" ></iframe>
     <input type="hidden" name="filename">
      </td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="ȷ��"></td>
	</tr>
</table>
</form>
<%  end if 
  if  action="save" then
     dim nettitle,netbody,workfile,filename
	     nettitle=request.form("nettitle")
		 netbody=request.form("netbody")
		 workfile=request.form("filename")
		 filename=request.form("workfile")
		 if nettitle="" or netbody="" then
		 response.write"<script>alert('�뷵�ذ�������д��ϸ');history.back();</script>"
             else
			 netbody=server.HTMLEncode(netbody)
			 sql="insert into worknet(nettitle,netbody,nettime,filename,workfile) values('"&nettitle&"','"&netbody&"','"&now()&"','"&filename&"','"&workfile&"')"
		         conn.execute(sql)
		 %>
		 
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B4E7F1">��ӳɹ�</th>
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