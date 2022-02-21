<!--#include file="admintop.asp"-->
<!--#include file="code.asp"-->
<!--#include file="../mzfunc.asp"-->
<!--#include file="isadmin.asp"-->
<%
        dim action,id
	    id=request.querystring("id")
	 if id="" or isnull(id) then
	    response.write"<script>alert('数据丢失\n此页面不支持刷新');history.back();</script>"
	  end if
	     action=request.querystring("action")
	 if action="" or isnull(action) then
	    sql="select tjid,tjtitle,tjbody from mztj where tjid="&id
	    set rs=conn.execute(sql)
	 if not rs.eof then
%>
<form name="myform" method="post" action="edittj.asp?action=edit&id=<%=id%>">
  <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr bgcolor="#B1E1F3"> 
      <td height="30" colspan="2"> 
        <div align="center"><strong>修改统计</strong></div></td>
    </tr>
    <tr> 
      <td width="17%">统计标题:</td>
      <td width="83%"><input name="tjtitle" type="text" id="title" value="<%=rs("tjtitle")%>"></td>
    </tr>
    <tr bgcolor="#B1E1F3"> 
      <td height="25" colspan="2"> <div align="center"><strong>统计内容</strong></div></td>
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
          <textarea name="body" cols="70" rows="15" id="body"><%=rs("tjbody")%></textarea>
        </div></td>
    </tr>
    <tr> 
      <td colspan="2"> <div align="center"> 
          <input type="submit" name="ccc" value="修改">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type="reset" name="Submit2" value="取消">
        </div></td>
    </tr>
  </table>
</form>
<%  else
          response.write"<script>alert('你所要操作的对象未找到\n请返回重试');history.back();</script>" 
    end if
	end if
    if action="edit" then
	    dim tjtitle,body,ccc
	    tjtitle=request.Form("tjtitle")
	    body=request.Form("body")
	    ccc=request.Form("ccc")
	if tjtitle="" or body="" then
       response.write"<script>alert('请返回把资料填写详细');history.back();</script>" 
     end if
	 if ccc<>"修改" then
	    response.write"<script>alert('不要胡闹\n不要胡闹');history.back();</script>" 
      end if
	      sql="update mztj set tjtitle='"&tjtitle&"', tjbody='"&body&"', tjtime='"&now()&"' where tjid="&id	   
          conn.execute(sql)
   %>
<table width="350" border="0" align="center">
  <tr>
    <th height="30" bgcolor="#B1E1F3">修改成功</th>
  </tr>
  <tr>
    <td>请选择你要返回的页面</td>
  </tr>
  <tr>
    <td><div align="center"><a href="admin_tj.asp">【返回上一页面】</a> <a href="adminbody.asp">【返回其它操作】</a></div></td>
  </tr>
</table>
<% end if %>
</body>
</html>
