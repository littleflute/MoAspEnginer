<!--#include file="admintop.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
     dim action,id
	 id=request.QueryString("id")
	 action=request.querystring("action")
	 if action="" or isnull(action) then
	 sql="select * from worknet where id="&id
	    set rs=conn.execute(sql)
		if rs.eof then
		response.write"<script>alert('对法起,你把要操作的对象未找到');history.back();</scrpt>"
              else
	   %>
<form name="frmCtoy" method="post" action="editworknet1.asp?action=edit&id=<%=id%>" >
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000" bordercolorlight="#000000">
    <tr bgcolor="#B4E7F1"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>添加网上事务文件</strong></div></td></tr>
    <tr>
      <td width="22%">网上事务标题:</td>
      <td height="25"> <input name="nettitle" type="text" size="30" value="<%=rs("nettitle")%>"></td>
    <tr>
	<tr>
      <td width="22%">内容:</td>
      <td width="78%"> 
        <textarea name="netbody" cols="60" rows="20"><%=rs("netbody")%></textarea>
      </td>
    </tr>
	<tr><td>附件名称:</td><td><input name="workfile" type="text" size="30" value="<%=rs("filename")%>"></td></tr>
 <tr> 
      <td height="40">上传的附件：</td>
      <td height="25"> 
     <input type="text" name="filename" value="<%=rs("workfile")%>">
        注意,如果你不是太熟悉的话.此处最好不要修改.</td>
  </tr>
  <tr>
      <td height="25">&nbsp;</td>
    <td><input type="submit" name="ups" value="确定"></td>
	</tr>
</table>
</form>
 <%  end if  
end if 
  if  action="edit" then
     dim nettitle,netbody,workfile,filename
	     nettitle=request.form("nettitle")
		 netbody=request.form("netbody")
		 workfile=request.form("filename")
		 filename=request.form("workfile")
		 if nettitle="" or netbody="" then
		 response.write"<script>alert('请返回把资料填写详细');history.back();</script>"
             else
			 netbody=server.HTMLEncode(netbody)
		     sql="update worknet set nettitle='"&nettitle&"',netbody='"&netbody&"',workfile='"&workfile&"',filename='"&filename&"',nettime='"&now()&"' where id="&id
		     conn.execute(sql)
		 %>
		 
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B4E7F1">修改成功</th>
  </tr>
  <tr>
    <td>你可以返回到上一页面查看</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_worknet.asp">【返回】</a></div></td>
  </tr>
</table>
<% end if
   end if %>
</body>
</html>