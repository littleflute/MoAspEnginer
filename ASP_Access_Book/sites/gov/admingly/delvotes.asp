<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id,action,errorcount
	errorcount=0
	id=trim(request.QueryString("id"))
	action=trim(request.querystring("action"))
	if id="" or action="" then
	response.write"<script>alert('�Բ���,���������ݶ�ʧ\n�뷵������');history.back();</script>"
       errorcount=1
	end if
	if action="del" and  errorcount=0 then
	sql="delete * from votes where voteid="&id
	conn.execute(sql)
%>
<table width="100%" border="0" align="center">
  <tr> 
    <th height="30" bgcolor="#B1E1F3"> <div align="center"><strong>ɾ���ɹ�</strong></div></th>
  </tr>
  <tr> 
    <td>�ļ��Ѿ�ɾ��,ɾ�������ݽ����ö�ʧ,����ʱ��С�Ĵ���.�뷵����һҳ���ˢ��ҳ����в鿴</td>
  </tr>
</table>
<% end if
  if action="edit" then
    sql="select voteid,votetitle from votes where voteid="&id
	set rs=conn.execute(sql)
	if not rs.eof then
  %>
  <form name="form1" method="post" action="delvotes.asp?id=<%=rs("voteid")%>&action=xiugai">
<table width="350" border="0" align="center">
    <tr bgcolor="#B1E1F3"> 
      <th height="30" colspan="2"> <div align="center"><strong>�༭ͶƱѡ��</strong></div></th>
  </tr>
  <tr> 
    <td width="74">�޸�:</td>
    <td width="266">
        <input type="text" name="votetitle" value="<%=rs("votetitle")%>">
      </td>
  </tr>
  <tr> 
    <td colspan="2"><div align="center">
          <input type="submit" name="Submi" value="�޸�">
        </div></td>
  </tr>
</table>
</form>
<% else
response.write"<script>alert('����Ҫ�����Ķ��󲻴���,�뷵������');history.bakc();</script>"
  end if
  end if
  if action="xiugai" then
  dim votetitle,Submi
     votetitle=trim(request.form("votetitle"))
	 Submi=request.form("Submi")
	 if votetitle="" or Submi<>"�޸�" then
response.write"<script>alert('�Ƿ�����,����δ֪����00ecaX7ae501.�뷵������');history.bakc();</script>"
          else
		  sql="update votes set votetitle='"&votetitle&"' where voteid="&id
		  conn.execute(sql)
		  end if
%>
<table width="100%" border="0" align="center">
  <tr> 
    <th height="30" bgcolor="#B1E1F3"> <div align="center"><strong>�޸ĳɹ�</strong></div></th>
  </tr>
  <tr> 
    <td>�ļ��Ѿ��޸�,�޸�ǰ�����ݽ����ö�ʧ,����ʱ��С�Ĵ���.�뷵����һҳ���ˢ��ҳ����в鿴</td>
  </tr>
</table>
<% end if %>
</body>
</html>
