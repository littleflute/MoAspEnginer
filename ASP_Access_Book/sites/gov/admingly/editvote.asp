<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim id,action
   action=request.querystring("action")
   id=request.querystring("id")
   if id="" or isnull(id) then
   response.write"<script>alert('�Բ���,���ݶ�ʧ\n��ҳ�治֧��ˢ��');history.back();</script>"
       end if
	   if action="" then
%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr bgcolor="#B1E1F3"> 
    <td height="30" colspan="2"> 
      <div align="center"><strong>��ѡ�����</strong></div></td>
  </tr>
  <tr> 
    <td width="53%"><a href="editvote.asp?id=<%=id%>&action=stat">�޸�ͶƱ������</a></td>
    <td width="47%"><a href="editvote.asp?id=<%=id%>&action=votes">�޸�ͶƱѡ����Ŀ</a></td>
  </tr>
  <tr>
    <td colspan="2"><div align="right"><a href="javascript:history.go(-1)">�����ء�</a></div></td>
  </tr>
</table>
<% end if
	 if action="stat" then
	  sql="select title,id from stat where id="&id
	       set rs=conn.execute(sql)
		   if rs.eof then
	 response.write"<script>alert('�Բ���,���ݶ�ʧ\n������Ҫ�Ķ���δ�ҵ�');history.back();</script>"
            else
		   %>
		   <form name="form1" method="post" action="editvote.asp?id=<%=rs("id")%>&action=xiugai">
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>���޸�ͶƱ������</strong></div></td>
  </tr>
  <tr> 
    <td>������:</td>
    <td>
        <input type="text" name="stat" value="<%=rs("title")%>">
     </td>
  </tr>
  <tr> 
    <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="�޸�">
        </div></td>
  </tr>
</table>
 </form>
 <% end if
     end if
  if action="votes" then
 %>
<table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" colspan="3" bgcolor="#B1E1F3"> 
      <div align="center"><strong>��ѡ����Ҫ�����Ķ�����б༭</strong></div></td>
  </tr>
  <tr> 
    <td width="76%"><div align="center">ͶƱѡ��</div></td>
    <td colspan="2"><div align="center">����</div></td>
  </tr>
  <%
      sql="select voteid,votetitle from votes where fromid="&id
	  set rs=conn.execute(sql)
	  if not rs.eof then
	  do while not rs.eof 
  %>
 <tr><td><div align="center"><%=rs("votetitle")%></div></td>
    <td width="13%"><div align="center"><a href="delvotes.asp?id=<%=rs("voteid")%>&action=edit">�޸�</a></div></td>
    <td width="11%"><div align="center"><a href="delvotes.asp?id=<%=rs("voteid")%>&action=del">ɾ��</a></div></td>
  </tr>
<%  rs.movenext
     loop
	 end if 
	 end if
    if action="xiugai" then
	dim stat
	stat=request.form("stat")
	sql="update stat set title='"&stat&"' where id="&id
	conn.execute(sql)
%>
<tr>
    <td><div align="center"><a href="addvotes.asp">����˴����ͶƱ</a></div></td>
  </tr>
</table>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>�޸ĳɹ�</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">�뷵����һҳ����в鿴</div></td>
  </tr>
  <tr> 
    <td><div align="right"><a href="admin_vote.asp">�����ء�</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
