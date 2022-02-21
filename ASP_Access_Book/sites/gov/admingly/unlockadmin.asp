<!--#include file="admintop.asp"-->
<!--#include file="../mzconst.asp"-->
<!--#include file="isadmin.asp"-->
<%    dim id,action
        id=request.querystring("id")
		action=request.querystring("action")
		if action="" or isempty(action) then
   sql="select * from admingly where islock>="&errorpwd
          set rs=conn.execute(sql)
             if not rs.eof then
response.write"<table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#000000'>"&_
"<tr align='center' bgcolor='#B1E1F3'> <td height='25' colspan='4'>'管理员状态'</td>"&_
 "</tr><tr><td>管理员ID</td><td>管理员姓名</td><td>状态</td><td>操作</td></tr>"
         do while not rs.eof
response.write"<tr><td>"&rs("user_name")&"</td><td>"&rs("names")&"</td><td>已锁定</td><td><a href='unlockadmin.asp?action=unlock&id="&rs("id")&"'>解锁</a></td></tr>"
        rs.movenext
		loop
		else 
		response.write"<tr align='center' bgcolor='#B1E1F3'> <td height='25' colspan='4'>暂无锁定</td></table>"
         end if
      response.write"</table>"
	  	else
     if action="unlock" then
	  if id="" or id=0 or isnumeric(id)<>true then
	     response.redirect"error.htm?errmsg='"&now()&"'"
		 else
		 sql="update admingly set islock=0 where id="&id
		 conn.execute(sql)
		 end if
	 response.write"<table width='300' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor='#000000'>"&_
	 "<tr><td height='25' align='center' bgcolor='#B1E1F3'>修改成功</td></tr>"&_
	 "<tr><td>已成功解除该管理员账号的锁定,你可以返回查看</td></tr><tr><td align='center'> <a href='unlockadmin.asp'>【返回】</a></td></tr></table>"&_
	 "</body></html>"
	 end if
	 end if
	 %>

  
  

