<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    DIM id,action
	action=request.querystring("action")
	id=request.QueryString("id")
	if id="" or isnull(id) then
	response.write"<script>alert('��������ʧ,\n�뷵������');history.back();</script>"
       end if
	   if action="" or isnull(action) then
	sql="select * from contact where id="&id
      set rs=conn.execute(sql)
	   if rs.eof then
response.write"<script>alert('����Ҫ�����Ķ���δ�ҵ�,\n�뷵������');history.back();</script>"
            else  
%>
<form name="form1" method="post" action="replyconsul.asp?id=<%=id%>&action=reply">
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td width="87">��ѯ����:</td>
      <td width="253"> <textarea name="content" rows="5"><%=rs("guestcontent")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td>�ظ�:</td>
      <td><textarea name="reply" rows="5"><%=rs("reply")%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="�� ��">
          �� 
          <input type="reset" name="Submit2" value="ȡ ��">
        </div></td>
    </tr>
  </table>
</form>
<% end if
   end if
   if action="reply" then
     dim reply,content
	 reply=trim(request.Form("reply"))
	 content=trim(request.Form("content"))
	 if reply="" or content="" then
	 response.write"<script>alert('���������д��ϸ�ٱ���');history.back();</script>"
	         else
			 sql="update contact set guestcontent='"&content&"',reply='"&reply&"' where id="&id
			    conn.execute(sql)
				end if
			 %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="25" bgcolor="#B1E1F3"> 
      <div align="center"><strong>�༭�ظ��������</strong></div></td>
  </tr>
  <tr>
    <td>����Է�����һҳ���в鿴,Ϊ��������ѯ���뾡����Ҫ������������</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_consul.asp">�����ء�</a></div></td>
  </tr>
</table>
<% end if%>
</body>
</html>
