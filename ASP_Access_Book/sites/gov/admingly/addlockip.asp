<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="../mzfunc.asp"-->
<%
     dim  action
	 action=request.querystring("action")
	 if action="" or isnull(action) then
%>
<form name="form1" method="post" action="addlockip.asp?action=add">
  <table width="300" border="1" align="center" bordercolor="#000000">
    <tr align="center" bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"><strong>添加锁定ip</strong></td>
  </tr>
  <tr> 
    <td width="81">IP地址:</td>
    <td width="209"><input name="badip" type="text" maxlength="15"></td>
  </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type="submit" name="Submit" value="添 加">
      </td>
  </tr>
</table>
     </form>
	 <%
	   end if
	   if action="add" then
	   dim lockbadip
	   lockbadip=request.Form("badip")
	   if lockbadip="" or isnull(lockbadip) then
	   response.write"<script>alert('对不起，添加的ip地址不能为空');histroy.back();</script>"
	     else
		 if islock(lockbadip) then
		 sql="insert into killip(badip,times) values('"&lockbadip&"','"&now()&"')"
	        conn.execute(sql)
	          else
			  response.write"<script>alert('对不起,请输入合法的ip地址');history.back();</script>"
			  end if
	 %>
<table width="350" border="1" align="center" bordercolor="#000000">
  <tr>
    <td height="25" align="center" bgcolor="#B1E1F3"><strong>添加成功</strong></td>
  </tr>
  <tr>
    <td>你可以返回查看你所添加的项</td>
  </tr>
  <tr>
    <td align="center"><a href="unlock.asp">【返回】</a></td>
  </tr>
</table>
<% 
   end if
   end if
%>
</body>
</html>
