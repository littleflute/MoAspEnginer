<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%      dim action,webaddr,fnames
         action=request.querystring("action")
        if action="add" then
	     if request.form("Sumit")<>"添 加" then
		 response.write"<script>alert('不要胡闹\n返回重试');history.back();</script>" 
                 end if
		 fnames=request.form("fnames")
		 webaddr=request.form("webaddr")
		 sql="select * from aboutlink where webaddr='"&webaddr&"' or zwname='"&fnames&"'"
		      set rs=conn.execute(sql)
			  if not rs.eof then
			 response.write"<script>alert('此单位已经添加了\n请返回');history.back();</script>" 
		         end if
			 if webpage(webaddr) then
		 sql="insert into aboutlink(webaddr,zwname) values('"&webaddr&"','"&fnames&"')"
               conn.execute(sql)
			   %>			   
<table width="100%" border="0" align="center">
  <tr>
    <th>添加成功</th>
  </tr>
  <tr>
    <td>数据已经添加,请返回首页进行查看.或选择你要返回的地点.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_link.asp">【返回】</a><a href="adminbody.asp">【返回前一页】</a> 
        </div></td>
  </tr>
</table>
<%
          else
			  response.write"<script>alert('对不起,网址不合法\n请返回重新填写');history.back();</script>"
			    end if
				end if
			    if action="" or isnull(action) then
			   %>
<form name="form1" method="post" action="addlink.asp?action=add">			   
  <table width="350" border="0" align="center">
    <tr> 
    <th colspan="2">添加链接</th>
  </tr>
  <tr> 
    <td>链接的单位名称:</td>
    <td>
        <input type="text" name="fnames">
      </td>
  </tr>
  <tr> 
    <td>链接单位的网址:</td>
    <td><input type="text" name="webaddr"></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
    <td><input type="submit" name="Sumit" value="添 加"></td>
  </tr>
</table>
</form>
<% end if %>
</body>
</html>
