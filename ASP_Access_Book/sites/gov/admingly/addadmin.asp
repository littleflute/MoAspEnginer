<!--#include file="admintop.asp"-->
<!--#include file="adminMd5.asp"-->
<!--#include file="isadmin.asp"-->
  <%
      dim action
	  action=request.querystring("action")
	  if action="" or isnull(action) then 
   %>
<form name="form1" method="post" action="addadmin.asp?action=save">
  <table width="300" border="1" align="center" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> <div align="center"><strong>��ӹ���Ա</strong></div></td>
    </tr>
    <tr> 
      <td width="78">����ԱID:</td>
      <td width="212"> <input type="text" name="username"> </td>
    </tr>
    <tr> 
      <td>����:</td>
      <td><input type="password" name="password"></td>
    </tr>
   <tr> 
      <td>ȷ������:</td>
      <td><input type="password" name="password1"></td>
    </tr>
    <tr> 
      <td>��ʵ����:</td>
      <td><input type="text" name="names"></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="ȷ��">
          �� 
          <input type="reset" name="Submit2" value="����">
        </div></td>
    </tr>
  </table>
</form>
<% 
   else 
   if action="save" then
   dim username,password,names,errorcount
   errorcount=0
   username=trim(request.form("username"))
   password=request.form("password")
   password1=request.form("password1")
   if password<>password1 then
   response.write"<script>alert('�Բ���,������������벻һ��');history.back();</script>"
     errorcount=1
	 end if
   names=request.form("names")
 if username="" or password="" or names="" then
 response.write"<script>alert('�Բ���,�������������д��ϸ');history.back();</script>"
   errorcount=1
    else
	if password<>trim(password)  then
	  response.write"<script>alert('�������벻Ҫ�����ո�������Ƿ��ַ�\n���ұ�����6λ����');history.back();</script>"
          errorcount=1
		 else
		 sql="select * from admingly where user_name='"&username&"'"
		      set rs=conn.execute(sql)
			  if not rs.eof then
			  errorcount=1
	  response.write"<script>alert('�˹���Ա�Ѿ�����\n�뷵������');history.back();</script>"
		        end if
		    end if
	       end if
	if errorcount=0 then
	    password=md5(password)
       Set rs = Server.CreateObject("ADODB.RecordSet")
	   sql="select * from admingly where id is null" 
	   rs.open sql,conn,3,3
	   rs.addnew
	   rs("user_name")=username
	   rs("password")=password
	   rs("islock")=0
	   rs("names")=names
	   rs.update 
	%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="25" bgcolor="#B1E1F3"> <div align="center"><strong>��ӳɹ�</strong></div></td>
  </tr>
  <tr> 
    <td>����Է�����һҳ���в鿴,������ӵĹ���Ա���к���һ����Ȩ��,�벻Ҫ������ӹ���Ա.</td>
  </tr>
  <tr> 
    <td><div align="center"><a href="adminbody.asp">�����ء�</a></div></td>
  </tr>
</table>
<%
   end if
   end if
   end if
%>
</body>
</html>
