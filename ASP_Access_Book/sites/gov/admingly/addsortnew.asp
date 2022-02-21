<!--#include file="code.asp"-->
<!--#include file="admintop.asp"-->
<!--#include file="isadmin.asp"-->
<form name="myform" method="post" action="savesortnew.asp">
  <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td height="30" colspan="2" bgcolor="#B1E1F3"> 
        <div align="center"><strong>添加栏目新闻</strong></div></td>
  </tr>
  <tr> 
    <td>选择栏目分类:</td>
    <td>
        <select name="sort">
          <option value="" selected>请选择分类</option>
          <%  dim sorts,i
		sql="select * from allsort order by id" 
		   set rs=conn.execute(sql)
		     sorts=rs.getrows
			 for i=0 to ubound(sorts,2)
		%>
          <option value="<%=sorts(0,i)%>"><%=sorts(1,i)%></option>
          <% next %>
        </select>
      </td>
  </tr>
  <tr> 
      <td>栏目新闻标题:</td>
    <td><input name="title" type="text" id="title"></td>
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
          <textarea name="body" cols="70" rows="15" id="body"></textarea>
        </div></td>
  </tr>
  <tr>
      <td colspan="2"> <div align="center">
          <input type="submit" name="ccc" value="提交">
          &nbsp;&nbsp;&nbsp;&nbsp;<input type="reset" name="Submit2" value="取消">
        </div></td></tr>
</table>
</form>
</body>
</html>
