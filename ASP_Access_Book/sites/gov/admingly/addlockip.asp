<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="../mzfunc.asp"-->
<%
     dim  action
	 action=request.querystring("action")
	 if action="" or isnull(action) then
%>
<form name="form1" method="post" action="addlockip.asp?action=add">
  <table width="300" border="1" align="center" bordercolor="#000000">
    <tr align="center" bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"><strong>�������ip</strong></td>
  </tr>
  <tr> 
    <td width="81">IP��ַ:</td>
    <td width="209"><input name="badip" type="text" maxlength="15"></td>
  </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type="submit" name="Submit" value="�� ��">
      </td>
  </tr>
</table>
     </form>
	 <%
	   end if
	   if action="add" then
	   dim lockbadip
	   lockbadip=request.Form("badip")
	   if lockbadip="" or isnull(lockbadip) then
	   response.write"<script>alert('�Բ�����ӵ�ip��ַ����Ϊ��');histroy.back();</script>"
	     else
		 if islock(lockbadip) then
		 sql="insert into killip(badip,times) values('"&lockbadip&"','"&now()&"')"
	        conn.execute(sql)
	          else
			  response.write"<script>alert('�Բ���,������Ϸ���ip��ַ');history.back();</script>"
			  end if
	 %>
<table width="350" border="1" align="center" bordercolor="#000000">
  <tr>
    <td height="25" align="center" bgcolor="#B1E1F3"><strong>��ӳɹ�</strong></td>
  </tr>
  <tr>
    <td>����Է��ز鿴������ӵ���</td>
  </tr>
  <tr>
    <td align="center"><a href="unlock.asp">�����ء�</a></td>
  </tr>
</table>
<% 
   end if
   end if
%>
</body>
</html>
