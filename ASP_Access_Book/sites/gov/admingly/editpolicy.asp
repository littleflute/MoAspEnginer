<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim id
	 id=request.querystring("id")
	 action=request.querystring("action")
	 if action="" or isnull(action) then
	 if id="" or isnull(id) then
	 response.write"<script>alert('��������\n�뷵������.');history.back();</script>"
	      else
	 sql="select * from policy where id="&id
	 set rs=conn.execute(sql)
         end if
		 if rs.eof then
		 response.write"<script>alert('��ȷ�����������Ķ���ȷʵ����\n�뷵������.');history.back();</script>"
           else
%>
<form name="frmCtoy" method="post" action="editpolicy.asp?action=edit&id=<%=id%>" >
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B4E7F1"> 
        <div align="center"><strong>�༭���߷���</strong></div></td></tr>
    <tr>
      <td width="22%">���߷������:</td>
      <td height="25"> <input name="pytitle" type="text" size="30" value="<%=rs("pytitle")%>"></td>
    <tr>
	<tr>
      <td width="22%">���߷�������:</td>
      <td width="78%"> 
        <textarea name="pybody" cols="60" rows="20"><%=rs("pybody")%></textarea>
      </td>
    </tr>
 <tr> 
      <td height="40">�ϴ�ͼƬ��</td>
      <td height="25"> 
     <input type="text" name="filename" value="<%=rs("pyimg")%>">
        ���ͼƬ�ļ�������̫���,�벻Ҫ���޸�</td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="�޸�"></td>
	</tr>
</table>
</form>
<% end if
  elseif action="edit" then
  dim pytitle,pybody,pyimg,ups,errorcount
       errorcount=""
    pytitle=trim(request.form("pytitle"))
	pybody=trim(server.htmlencode(request.form("pybody")))
	pyimg=trim(request.form("filename"))
	ups=request.form("ups")
	if ups<>"�޸�" then
	response.write"<script>alert('�벻Ҫ���Ҳ���\nлл');history.back();</script>"
        errorcount="������"
		end if
		if pytitle="" or pybody="" then
	response.write"<script>alert('���������д��ϸ\nлл');history.back();</script>"
         errorcount="������"
		  end if
		  if pyimg<>"" then
         if not isexists(pyimg) then
	response.write"<script>alert('��༭���ļ�������,\n�뷵������');history.back();</script>"
               errorcount="������"
			    end if
				end if
         if errorcount="" then
		 sql="update policy set pytitle='"&pytitle&"', pybody='"&pybody&"', times='"&now()&"', pyimg='"&pyimg&"' where id="&id&"" 
             conn.execute(sql)
   %>
   <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>�޸ĳɹ�</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">�޸ĺ���ļ�������ҳ�����޸ĺ��������ʾ,����ʱ��һ�������Դ�.�����У,��ѡ���㷵��Ҫ���ص�ҳ��.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="admin_policy.asp">�����ء�</a><a href="adminbody.asp">����������������</a></div></td>
  </tr>
</table>
<% end if
   end if %>
</body>
</html>