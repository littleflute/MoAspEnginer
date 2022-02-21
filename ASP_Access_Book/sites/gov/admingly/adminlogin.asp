<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="adminMd5.asp"-->
<% 
	action=request.querystring("action")
	if action="" or isnull(action) then
%>
<form name="form1" method="post" action="adminlogin.asp?action=yz">
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>管理员登陆</strong></div></td>
    </tr>
    <tr background="topBar_bg.gif"> 
      <td width="128"> 
        <div align="center"><font color="#FF0000">管理员账号:</font></div></td>
      <td width="212"> 
        <input type="text" name="username"> </td>
    </tr>
    <tr background="topBar_bg.gif"> 
      <td> 
        <div align="center"><font color="#FF0000">密码:</font></div></td>
      <td> 
        <input type="password" name="password"></td>
    </tr>
    <tr background="topBar_bg.gif"> 
      <td colspan="2"> 
        <div align="center">
          <input type="submit" name="cdf" value="确 定">
          　 
          <input type="reset" name="Submit2" value="取 消">
        </div></td>
    </tr>
  </table>
 </form>
 <%  else if action="yz" then 
       dim username,password,cdf
	   username=request.form("username")
	   password=request.form("password")
	   cdf=request.form("cdf")
	   if cdf<>"确 定" then
	   response.write"<script>alert('对不起,\n不要胡闹不要胡闹');window.location.href='adminlogin.asp';</script>"
           end if
		   sql="select user_name,password,islock from admingly where user_name='"&username&"'"
	       set rs=conn.execute(sql)
		     if rs.eof then
			 response.write"<script>alert('对不起,此账号不存在');window.location.href='adminlogin.asp'</script>"
			     else
				    if errorpwd<=rs("islock") then
					response.write"<script>alert('对不起，账号已锁定');window.location.href='adminlogin.asp';</script>"
				          else
				 if md5(password)<>rs("password") then
				   sql="update admingly set islock=islock+1 where user_name='"&username&"'"
			               conn.execute(sql)
						   errorpwd=errorpwd-rs("islock")
						   if errorpwd<=0 then
							 sql="insert into killip(username,badip,times) values('"&username&"','"&getip()&"','"&now()&"')"
							        conn.execute(sql)
							response.write"<script>alert('对不起,账号已被锁定');window.location.href='adminlogin.asp';</script>"	
								 else
				response.write"<script>alert('对不起,密码不正确,\n你还有"&errorpwd&"次机会');window.location.href='adminlogin.asp';</script>"
			           end if
					  else
					  session("username")=username
					  session("ispass")="pass"
					  sql="update admingly set islock=0 where user_name='"&username&"'"
					    conn.execute(sql)
						  if iscookies then
						response.Cookies("mzcoks")("unm")=username
						response.Cookies("mzcoks")("upwd")=password
						response.Cookies("mzcoks").expires=date+100
						  end if
					    response.Redirect"adminbody.asp"
					   end if
					   end if
					   end if
					   end if
					   end if
			 %>
</body>
</html>
