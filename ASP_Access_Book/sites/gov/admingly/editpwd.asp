<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<!--#include file="adminmd5.asp"-->
<%
     dim action
	 action=request.QueryString("action")
	 if action="" or isnull(action)  then
	 sql="select * from admingly where user_name='"&session("username")&"'"
         set rs=conn.execute(sql)
%>
 <form name="form1" method="post" action="editpwd.asp?action=edit&id=<%=rs("id")%>">
  <table width="350" border="1" align="center" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2" align="center"><strong>更改密码</strong></td>
    </tr>
    <tr> 
      <td width="101">原密码:</td>
      <td width="239"> <input type="password" name="ordpwd"> </td>
    </tr>
    <tr> 
      <td>新密码:</td>
      <td><input type="password" name="newpwd"></td>
    </tr>
    <tr> 
      <td>确认密码:</td>
      <td><input type="password" name="newpwd1"></td>
    </tr>
    <tr> 
      <td>真实姓名:</td>
      <td><input type="text" name="names" value="<%=rs("names")%>"></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"> 
        <input type="submit" name="Submit" value="确 定">
        　
<input type="reset" name="Submit2" value="取 消">
      </td>
    </tr>
  </table>
</form>
<%
  end if
  if action="edit" then
  dim id,ordpwd,newpwd,newpwd1,names,errorcount,i
  errorcount=0
  id=request.QueryString("id")
  if id="" or isnull(id) then
  response.write"<script>alert('对不起,操作数据丢失\n请返回重试');history.back();</script>"
     errorcount=1
	 else
     ordpwd=request.form("ordpwd")
	 newpwd=request.form("newpwd")
	 newpwd1=request.form("newpwd1")
	 names=request.form("names")
	 if ordpwd="" or newpwd="" or newpwd1="" or names="" then
    response.write"<script>alert('请把资料填写详细\n请返回重试');history.back();</script>"
	    errorcount=1
	       else
		   if newpwd<>newpwd1 then
    response.write"<script>alert('两次输入的密码不一致\n请返回重试');history.back();</script>"
	        errorcount=1
			end if
			end if
	     if newpwd<>trim(newpwd) then
   response.write"<script>alert('密码中请不要包含空格逗号及其它非法字符');history.back();</script>"
	          errorcount=1
			  end if
			for i=1 to len(newpwd)-1
			if asc(mid(newpwd,i,1))>255 or asc(mid(newpwd,i,1))<0 then
  response.write"<script>alert('密码中请不要包含空格逗号及其它非法字符');history.back();</script>"
	                     exit for
						 errorcount=1
						 end if
	                       next
						   end if
      sql="select * from admingly where user_name='"&session("username")&"'"
	  Set rs = Server.CreateObject("ADODB.RecordSet")
	    rs.open sql,conn,1,3
		if not rs.eof then
		  if md5(ordpwd)<>rs("password") then
 response.write"<script>alert('原密码不正确\n请反回重新输入');history.back();</script>"
	             errorcount=1
				 end if
	         else
 response.write"<script>alert('操作数据丢失\n请反回重新输入');history.back();</script>"
			errorcount=1
			end if 			 
	        if errorcount=0  then
			rs("password")=md5(newpwd)
			rs("names")=names
			rs.update
			session("ispass")=""
			%>
			<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="30" align="center" bgcolor="#B4E7F1"><strong>修改成功</strong></td>
  </tr>
  <tr>
    <td>你可以返回到上一页面查看</td>
  </tr>
  <tr>
    <td align="center"> <a href="adminbody.asp">【返回】</a></td>
  </tr>
</table>
<%end if
  end if %>
</body>
</html>
