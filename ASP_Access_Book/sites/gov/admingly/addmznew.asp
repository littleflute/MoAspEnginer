<!--#include file="admintop.asp"-->
<!--#include file="code.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
   dim action
       action=request.querystring("action")
	   if action="" or isnull(action) then
%>
<form name="myform" method="post" action="addmznew.asp?action=add">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> <div align="center"><strong>添加栏目新闻</strong></div></td>
    </tr>
    <tr> 
      <td width="20%">栏目新闻标题:</td>
      <td width="80%"><input name="title" type="text" id="title"></td>
    </tr>
	<tr><td>发布单位:</td><td><input name="writer" type="text" value="政府网"></td></tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> 
        <div align="center"><strong>新闻内容</strong></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><img onClick=bold() src="pic/icon_editor_bold.gif" width="22" height="22" alt="粗体" border="0"> 
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
          <img onClick=shadow() height=22 alt=阴影字 src="pic/shadow.gif" width=23 border=0></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> 
          <textarea name="body" cols="70" rows="15" id="body"></textarea>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="ccc" value="提交">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="取消">
        </div></td>
    </tr>
  </table>
</form>
<% end if
    if action="add" then
	    dim nstitle,nsbody,writers,ccc,errorcount
		    errorcount=0
	        nstitle=trim(request.form("title"))
	        nsbody=trim(request.form("body"))
	        writers=trim(request.form("writer"))
	        ccc=request.form("ccc")
	    if nstitle="" or nsbody="" or writers="" then
	       response.write"<script>alert('对不起,请返回把资料填写详细');history.back();</script>"
           errorcount=1
	    else
		    if ccc<>"提交" then
	           response.write"<script>alert('对不起,发生未知错误');history.back();</script>"
                errorcount=1
			 end if
	   end if
    if errorcount=0 then
	   nsbody=server.htmlencode(nsbody)
	   sql="insert into news(nsbody,nstitle,times,writers,clicks) values('"&nsbody&"','"&nstitle&"','"&now()&"','"&writers&"',0 )"     
       conn.execute(sql)
%>
<table width="350" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
  <tr>
    <th height="30" bgcolor="#B1E1F3">添加成功</th>
  </tr>
  <tr>
    <td>数据已成功添加,你可以返回查看.</td>
  </tr>
  <tr>
    <td><div align="center"><a href="javascript:window.close()">【关闭窗口】</a> </div></td>
  </tr>
</table>
<%   end if
  end if   %>
</body>
</html>
