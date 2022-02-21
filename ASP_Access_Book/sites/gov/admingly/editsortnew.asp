<!--#include file="code.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<%
    dim id,sortnew,action
	id=request.querystring("id")
	action=request.querystring("action")
	if id="" or isnull("id") then
	response.write"<script>alert('操作出现错误\n请返回重试');history.back();</script>"
         end if
  if action="" or isnull(action) then
   sql="select * from allarti where id="&id
    set rs=conn.execute(sql)
		  if rs.eof then
response.write"<script>alert('对不起,你所要操作的对象不存在\n请返回重试');history.back();</script>"
           else
		   sortnew=rs.getrows
%>
<form name="myform" method="post" action="editsortnew.asp?action=edit&id=<%=sortnew(0,0)%>">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> 
        <div align="center"><strong>添加栏目新闻</strong></div></td>
  </tr>
  <tr> 
    <td>选择栏目分类:</td>
    <td>
        <select name="sort">
          <%  dim sorts,i
		sql="select * from allsort order by id" 
		   set rs=conn.execute(sql)
		     sorts=rs.getrows
			 for i=0 to ubound(sorts,2)
		       if cint(sorts(0,i))=cint(sortnew(3,0)) then
		%>
        <option value="<%=sorts(0,i)%>" selected><%=sorts(1,i)%></option>
            <% else %>
          <option value="<%=sorts(0,i)%>"><%=sorts(1,i)%></option>
          <%
		   end if
		   next %>
        </select>
      </td>
  </tr>
  <tr> 
      <td>栏目新闻标题:</td>
    <td><input name="title" type="text" id="title" value="<%=sortnew(1,0)%>"></td>
  </tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> 
        <div align="center"><strong>栏目新闻内容</strong></div></td>
  </tr>
  <tr><td colspan="2"><div align="center"><img onClick=bold() src="pic/icon_editor_bold.gif" width="22" height="22" alt="粗体" border="0"> 
          <img onClick=italicize() src="pic/icon_editor_italicize.gif" width="23" height="22" alt="斜体" border="0"> 
          <img onClick=underline() src="pic/icon_editor_underline.gif" width="23" height="22" alt="下划线" border="0"> 
          <img onClick=center() src="pic/icon_editor_center.gif" width="22" height="22" alt="居中" border="0"> 
          <img onClick=hyperlink() src="pic/icon_editor_url.gif" width="22" height="22" alt="超级连接" border="0"> 
          <img onClick=email() src="pic/icon_editor_email.gif" width="23" height="22" alt="Email连接" border="0"> 
          <img onClick=image() src="pic/icon_editor_image.gif" width="23" height="22" alt="图片" border="0"> 
          <img onClick=flash() src="pic/swf.gif" width="23" height="22" alt="Flash图片" border="0"> 
          <img onClick=showcode() src="pic/icon_editor_code.gif" width="22" height="22" alt="编号" border="0"> 
          <img onClick=quote() src="pic/icon_editor_quote.gif" width="23" height="22" alt="引用" border="0"> 
          <img onClick=list() src="pic/icon_editor_list.gif" width="23" height="22" alt="目录" border="0"> 
          <img onClick=setfly() height=22 alt=飞行字 src="pic/fly.gif" width=23 border=0> 
          <img onClick=move() height=22 alt=移动字 src="pic/move.gif" width=23 border=0> 
          <img onClick=glow() height=22 alt=发光字 src="pic/glow.gif" width=23 border=0> 
          <img onClick=shadow() height=22 alt=阴影字 src="pic/shadow.gif" width=23 border=0></div></td></tr>
  <tr> 
    <td colspan="2"><div align="center">
          <textarea name="body" cols="70" rows="15" id="body"><%=sortnew(2,0)%></textarea>
        </div></td>
  </tr>
  <tr>
      <td colspan="2"> <div align="center">
          <input type="submit" name="ccc" value="修改">
          &nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="Submit2" value="取消">
        </div></td></tr>
</table>
</form>
<% end if 
  elseif action="edit" then
      frominto=request.form("sort")
	  title=request.form("title")
	  body=request.form("body")
	if frominto="" or title="" or body="" then
	response.write"<script>alert('对不起,请把资料填写详细后再修改');history.back();</script>"  
           else
		   sql="update allarti set frominto="&frominto&", title='"&title&"', body='"&body&"', times='"&now()&"' where id="&id
              conn.execute(sql)
            end if 
%>
<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr> 
    <td height="30" bgcolor="#B1E1F3"> <div align="center"><strong>修改成功</strong></div></td>
  </tr>
  <tr> 
    <td><div align="center">修改后的文件将在主页面上显示,操作时请一定谨慎对待.请选择你返回要返回的页面.</div></td>
  </tr>
  <tr> 
    <td height="20"> <div align="center"><a href="sort_new.asp">【返回】</a><a href="adminbody.asp">【继续其它操作】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
