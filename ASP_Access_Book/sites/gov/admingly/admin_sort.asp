<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<div align="center">
  <%
   dim id,action,sorts
   id=request.querystring("id")
   action=request.querystring("action")
Set rs = Server.CreateObject("ADODB.RecordSet")
	if action="" or isnull(action) then
	sql="select * from allsort order by id"
		rs.open sql,conn,1,3
	%>
  <table width="200" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> 
        <div align="center"><strong>������Ŀ</strong></div></td>
  </tr>
  <% do while not rs.eof %>
  <tr> 
    <td><%=rs("sort")%></td>
    <td><div align="center"><a href="admin_sort.asp?action=edit&id=<%=rs("id")%>">�޸�</a></div></td>
  </tr>
  <%  rs.movenext
     loop
  %>
  <tr> 
    <td colspan="2">&nbsp;&nbsp;&nbsp;��վһ����17����Ŀ,�������澭���˾��Ĳ������Ű�, �����ݲ��ṩ�������Ŀ�Ĺ���,���֮�����Ŀ������ҳ���ϲ�̫����.���Խ��������޸���ĿʱҲΪ����������ۿ���,ԭ�ȵ���Ŀ���ĸ��ֵ�,�޸ĺ���������Ҳ���ĸ���, 
      �����һЩ�ĸ���,��һЩ�����.����������������۴����ǳ����õ�Ӱ��.</td>
  </tr>
</table>
<%  end if 
 if action="edit" then
if id="" or isnull(id) then
response.write"<script>alert('�Բ���,�������ִ���,�뷵������');history.back();</script>"
   else
   sql="select * from allsort where id="&id
   rs.open sql,conn,1,1
   end if
   if not rs.eof then %>
   <form name="form1" method="post" action="admin_sort.asp?action=xiugai&id=<%=rs("id")%>">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
      <tr> 
        <td height="30" colspan="2" bgcolor="#B1E1F3"> 
          <div align="center"><strong>��Ŀ�����޸�</strong></div></td>
  </tr>
  <tr> 
    <td>�޸���Ŀ����:</td>
    <td>
        <input type="text" name="sort" value="<%=rs("sort")%>">
      </td>
  </tr>
  <tr> 
        <td colspan="2"> <div align="center">
            <input type="submit" name="Submit" value="�޸�">
          </div></td>
  </tr>
</table>
</form>
<% else  response.write"<script>alert('�Բ���,����Ҫ�����Ķ���δ�ҵ�,�뷵������');history.back();</script>"
   end if
  rs.close
   end if
%>
<%  if action="xiugai" then
      sorts=request.form("sort")
	  if sorts="" or isnull(sorts) then
  response.write"<script>alert('�Բ���,��Ŀ���Ʋ���Ϊ��,�뷵������');history.back();</script>"
       end if
	   sql="select * from allsort where id="&id
	   rs.open sql,conn,1,3
	   if not rs.eof then
	   rs("sort")=sorts
	   rs.update %>
	   
	   <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B1E1F3"> 
      <div align="center"><strong>�޸ĳɹ�</strong></div></td>
  </tr>
  <tr>
    <td>�����޸ĳɹ�,�����޸ĺ���������ݽ�ֱ��������������ʾ.����ʱ���������</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_sort.asp">�����ء�</a><a href="adminbody.asp">����������������</a></div></td>
  </tr>
</table>
	   <%else
	  response.write"<script>alert('�Բ���,����Ҫ�����Ķ���δ�ҵ�,�뷵������');history.back();</script>"
	   end if
 end if
%>
</div>
</body>
</html>
