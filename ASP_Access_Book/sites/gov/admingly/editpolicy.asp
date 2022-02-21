<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim id
	 id=request.querystring("id")
	 action=request.querystring("action")
	 if action="" or isnull(action) then
	 if id="" or isnull(id) then
	 response.write"<script>alert('操作错误\n请返回重试.');history.back();</script>"
	      else
	 sql="select * from policy where id="&id
	 set rs=conn.execute(sql)
         end if
		 if rs.eof then
		 response.write"<script>alert('请确认你所操作的对象确实存在\n请返回重试.');history.back();</script>"
           else
%>
<form name="frmCtoy" method="post" action="editpolicy.asp?action=edit&id=<%=id%>" >
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B4E7F1"> 
        <div align="center"><strong>编辑政策法规</strong></div></td></tr>
    <tr>
      <td width="22%">政策法规标题:</td>
      <td height="25"> <input name="pytitle" type="text" size="30" value="<%=rs("pytitle")%>"></td>
    <tr>
	<tr>
      <td width="22%">政策法规内容:</td>
      <td width="78%"> 
        <textarea name="pybody" cols="60" rows="20"><%=rs("pybody")%></textarea>
      </td>
    </tr>
 <tr> 
      <td height="40">上传图片：</td>
      <td height="25"> 
     <input type="text" name="filename" value="<%=rs("pyimg")%>">
        如果图片文件名您不太清楚,请不要乱修改</td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="修改"></td>
	</tr>
</table>
</form>
<% end if
  elseif action="edit" then
  dim pytitle,pybody,pyimg,ups,errorcount
       errorcount=""
    pytitle=trim(request.form("pytitle"))
	pybody=trim(server.htmlencode(request.form("pybody")))
	pyimg=trim(request.form("filename"))
	ups=request.form("ups")
	if ups<>"修改" then
	response.write"<script>alert('请不要胡乱操作\n谢谢');history.back();</script>"
        errorcount="出错了"
		end if
		if pytitle="" or pybody="" then
	response.write"<script>alert('请把资料填写详细\n谢谢');history.back();</script>"
         errorcount="出错了"
		  end if
		  if pyimg<>"" then
         if not isexists(pyimg) then
	response.write"<script>alert('你编辑的文件不存在,\n请返回重试');history.back();</script>"
               errorcount="出错了"
			    end if
				end if
         if errorcount="" then
		 sql="update policy set pytitle='"&pytitle&"', pybody='"&pybody&"', times='"&now()&"', pyimg='"&pyimg&"' where id="&id&"" 
             conn.execute(sql)
   %>
   <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>修改成功</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">修改后的文件将在主页上以修改后的内容显示,操作时请一定谨慎对待.认真对校,请选择你返回要返回的页面.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="admin_policy.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
<% end if
   end if %>
</body>
</html>