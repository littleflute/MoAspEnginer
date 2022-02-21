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
        <div align="center"><strong>所有栏目</strong></div></td>
  </tr>
  <% do while not rs.eof %>
  <tr> 
    <td><%=rs("sort")%></td>
    <td><div align="center"><a href="admin_sort.asp?action=edit&id=<%=rs("id")%>">修改</a></div></td>
  </tr>
  <%  rs.movenext
     loop
  %>
  <tr> 
    <td colspan="2">&nbsp;&nbsp;&nbsp;本站一共有17个栏目,在主界面经过了精心布置与排版, 所以暂不提供添加新栏目的功能,添加之后的栏目排在主页面上不太美观.所以建议你在修改栏目时也为主界面的美观考虑,原先的栏目是四个字的,修改后尽量保持他也是四个字, 
      如果有一些四个字,有一些五个字.这样会给主界面美观带来非常不好的影响.</td>
  </tr>
</table>
<%  end if 
 if action="edit" then
if id="" or isnull(id) then
response.write"<script>alert('对不起,操作出现错误,请返回重试');history.back();</script>"
   else
   sql="select * from allsort where id="&id
   rs.open sql,conn,1,1
   end if
   if not rs.eof then %>
   <form name="form1" method="post" action="admin_sort.asp?action=xiugai&id=<%=rs("id")%>">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
      <tr> 
        <td height="30" colspan="2" bgcolor="#B1E1F3"> 
          <div align="center"><strong>栏目名称修改</strong></div></td>
  </tr>
  <tr> 
    <td>修改栏目名称:</td>
    <td>
        <input type="text" name="sort" value="<%=rs("sort")%>">
      </td>
  </tr>
  <tr> 
        <td colspan="2"> <div align="center">
            <input type="submit" name="Submit" value="修改">
          </div></td>
  </tr>
</table>
</form>
<% else  response.write"<script>alert('对不起,您所要操作的对象未找到,请返回重试');history.back();</script>"
   end if
  rs.close
   end if
%>
<%  if action="xiugai" then
      sorts=request.form("sort")
	  if sorts="" or isnull(sorts) then
  response.write"<script>alert('对不起,栏目名称不能为空,请返回重试');history.back();</script>"
       end if
	   sql="select * from allsort where id="&id
	   rs.open sql,conn,1,3
	   if not rs.eof then
	   rs("sort")=sorts
	   rs.update %>
	   
	   <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" bgcolor="#B1E1F3"> 
      <div align="center"><strong>修改成功</strong></div></td>
  </tr>
  <tr>
    <td>数据修改成功,你所修改后的数据内容将直接在主界在上显示.操作时请谨慎处理</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_sort.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
	   <%else
	  response.write"<script>alert('对不起,你所要操作的对象未找到,请返回重试');history.back();</script>"
	   end if
 end if
%>
</div>
</body>
</html>
