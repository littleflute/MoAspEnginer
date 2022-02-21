<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    DIM id,action
	action=request.querystring("action")
	id=request.QueryString("id")
	if id="" or isnull(id) then
	response.write"<script>alert('操作符丢失,\n请返回重试');history.back();</script>"
       end if
	   if action="" or isnull(action) then
	sql="select * from contact where id="&id
      set rs=conn.execute(sql)
	   if rs.eof then
response.write"<script>alert('你所要操作的对象未找到,\n请返回重试');history.back();</script>"
            else  
%>
<form name="form1" method="post" action="replyconsul.asp?id=<%=id%>&action=reply">
  <table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td width="87">咨询内容:</td>
      <td width="253"> <textarea name="content" rows="5"><%=rs("guestcontent")%></textarea> 
      </td>
    </tr>
    <tr> 
      <td>回复:</td>
      <td><textarea name="reply" rows="5"><%=rs("reply")%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center">
          <input type="submit" name="Submit" value="提 交">
          　 
          <input type="reset" name="Submit2" value="取 消">
        </div></td>
    </tr>
  </table>
</form>
<% end if
   end if
   if action="reply" then
     dim reply,content
	 reply=trim(request.Form("reply"))
	 content=trim(request.Form("content"))
	 if reply="" or content="" then
	 response.write"<script>alert('请把资料填写详细再保存');history.back();</script>"
	         else
			 sql="update contact set guestcontent='"&content&"',reply='"&reply&"' where id="&id
			    conn.execute(sql)
				end if
			 %>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <td height="25" bgcolor="#B1E1F3"> 
      <div align="center"><strong>编辑回复动作完成</strong></div></td>
  </tr>
  <tr>
    <td>你可以返回上一页进行查看,为了尊重咨询者请尽量不要更改他的留言</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_consul.asp">【返回】</a></div></td>
  </tr>
</table>
<% end if%>
</body>
</html>
