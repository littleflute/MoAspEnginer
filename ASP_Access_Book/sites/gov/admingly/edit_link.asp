<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action,id,webaddr,fnames
	 action=request.querystring("action")
	 id=request.querystring("id")
	       if id="" or isnull(id) then
	 response.write"<script>alert('对不起\n操作出现了未知错误');history.back();</script>"
               end if
		   if action="" then
		   sql="select webaddr,zwname,id from aboutlink where id="&id
		   set rs=conn.execute(sql) %>
  <form name="form1" method="post" action="edit_link.asp?action=edit&id=<%=rs("id")%>">			   
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="#B1E1F3"> 
      <th height="30" colspan="2">编辑链接</th>
  </tr>
  <tr> 
    <td>链接的单位名称:</td>
    <td>
        <input name="fnames" type="text" value="<%=rs("zwname")%>" size="30">
      </td>
  </tr>
  <tr> 
    <td>链接单位的网址:</td>
    <td><input name="webaddr" type="text" value="<%=rs("webaddr")%>" size="30"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><input type="submit" name="Sumit" value="修 改"></td>
  </tr>
</table>
</form>
<% end if 
   if action="edit" then
   webaddr=trim(request.form("webaddr"))
   fnames=trim(request.form("fnames"))
   if webaddr="" or fnames="" then
   response.write"<script>alert('对不起,请将所需的数据输入完全');history.back();</script>"
        end if
		  if not webpage(webaddr) then
  response.write"<script>alert('对不起,网址不正确,请更改');history.back();</script>"
               else
   sql="update aboutlink set webaddr='"&webaddr&"', zwname='"&fnames&"' where id="&id			   
			   conn.execute(sql)
			  %>
			  
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">修改成功</th>
  </tr>
  <tr>
    <td>修改前的数据将永久性丢失,在修改时请谨慎处理</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_link.asp">【返回】</a></div></td>
  </tr>
</table>

			 <%  end if 
			     end if %>
</body>
</html>
