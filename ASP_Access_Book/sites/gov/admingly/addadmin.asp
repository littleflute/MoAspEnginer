<!--#include file="admintop.asp"-->
<!--#include file="adminMd5.asp"-->
<!--#include file="isadmin.asp"-->
  <%
      dim action
	  action=request.querystring("action")
	  if action="" or isnull(action) then 
   %>
<form name="form1" method="post" action="addadmin.asp?action=save">
  <table width="300" border="1" align="center" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> <div align="center"><strong>添加管理员</strong></div></td>
    </tr>
    <tr> 
      <td width="78">管理员ID:</td>
      <td width="212"> <input type="text" name="username"> </td>
    </tr>
    <tr> 
      <td>密码:</td>
      <td><input type="password" name="password"></td>
    </tr>
   <tr> 
      <td>确认密码:</td>
      <td><input type="password" name="password1"></td>
    </tr>
    <tr> 
      <td>真实姓名:</td>
      <td><input type="text" name="names"></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="确定">
          　 
          <input type="reset" name="Submit2" value="重置">
        </div></td>
    </tr>
  </table>
</form>
<% 
   else 
   if action="save" then
   dim username,password,names,errorcount
   errorcount=0
   username=trim(request.form("username"))
   password=request.form("password")
   password1=request.form("password1")
   if password<>password1 then
   response.write"<script>alert('对不起,你所输入的密码不一致');history.back();</script>"
     errorcount=1
	 end if
   names=request.form("names")
 if username="" or password="" or names="" then
 response.write"<script>alert('对不起,请把所有资料填写详细');history.back();</script>"
   errorcount=1
    else
	if password<>trim(password)  then
	  response.write"<script>alert('密码中请不要包含空格或其它非法字符\n并且必需是6位以上');history.back();</script>"
          errorcount=1
		 else
		 sql="select * from admingly where user_name='"&username&"'"
		      set rs=conn.execute(sql)
			  if not rs.eof then
			  errorcount=1
	  response.write"<script>alert('此管理员已经存在\n请返回重试');history.back();</script>"
		        end if
		    end if
	       end if
	if errorcount=0 then
	    password=md5(password)
       Set rs = Server.CreateObject("ADODB.RecordSet")
	   sql="select * from admingly where id is null" 
	   rs.open sql,conn,3,3
	   rs.addnew
	   rs("user_name")=username
	   rs("password")=password
	   rs("islock")=0
	   rs("names")=names
	   rs.update 
	%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="25" bgcolor="#B1E1F3"> <div align="center"><strong>添加成功</strong></div></td>
  </tr>
  <tr> 
    <td>你可以返回上一页进行查看,你所添加的管理员具有和你一样的权限,请不要随意添加管理员.</td>
  </tr>
  <tr> 
    <td><div align="center"><a href="adminbody.asp">【返回】</a></div></td>
  </tr>
</table>
<%
   end if
   end if
   end if
%>
</body>
</html>
