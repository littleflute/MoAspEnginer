<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%  dim id,action
     id=request.QueryString("id")
	 action=request.querystring("action")
	 if id="" or isnull(id) then
      response.write"<script>alert('�Բ���,��������ʧ\n�뷵������');history.back();</script>"
       end if
	if action="" or isnull(action) then
	 %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <th height="30" colspan="2" bgcolor="#B1E1F3">���������ļ��༭</th>
  </tr>
  <tr> 
    <td width="50%" height="21"><a href="editworknet1.asp?id=<%=id%>">����,���ݱ༭</a></td>
    <td width="50%"><a href="editworknet.asp?id=<%=id%>&action=acs">��������</a></td>
  </tr>
  <tr> 
    <td colspan="2"><div align="center"><a href="admin_worknet.asp">�����ء�</a></div></td>
  </tr>
</table>
   <% end if 
        if action="acs" then 
		dim filename,workfile
		%>
<table width="300" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <th height="30" colspan="2" bgcolor="#B1E1F3"> ����һ����</th>
  </tr>
  <%   sql="select filename,workfile from worknet where id="&id
            set rs=conn.execute(sql)
			if rs.eof then 
			response.write"<script>alert('�Բ���,����Ҫ�����Ķ���δ�ҵ�\n�뷵������');history.back();</script>"
			  end if
			if rs("filename")<>"" and rs("workfile")<>"" then
			filename=split(rs("filename"),";")
			workfile=split(rs("workfile"),";")
			 application.Lock
		    application("all")=ubound(filename)
			for i=0 to ubound(filename)
             application("filename"&i)=filename(i)
			 application("workfile"&i)=workfile(i)
			%>
  <tr> 
    <td width="208"><%=application("filename"&i)%></td>
    <td width="82"><a href="delworkfile.asp?id=<%=id%>&i=<%=i%>">ɾ��</a></td>
  </tr>
  <% next
     application.UnLock
     else 
	 response.write"<tr><td>û�и���</td></tr>"
	  end if
	 %>
  <tr> 
    <td colspan="2"><a href="addworkfile.asp?id=<%=id%>">����˴���Ӹ���</a></td>
  </tr>
</table>
<% end if
  %>
</body>
</html>
